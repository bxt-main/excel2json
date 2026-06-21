"""
excel2json.py
=============
通用 Excel → dict / JSON 解析器。

完全对齐 slot_simulator_1/excel2dict/excel2dict.py 的解析逻辑，
并做了如下改进：

  ① None 权重统一处理：权重为 None 时视为 0（原代码忽略此情况）
  ② 完整错误提示，字段名/行号/Sheet 名全部包含在异常信息中
  ③ 对外提供简洁接口：excel2dict（Excel→dict）和 excel2json（Excel→JSON文件）

## 核心数据格式约定

每个 Sheet 可以是两种格式之一：

  key-value 格式（默认）：
      每行 3 列：KEY  TYPE  VALUE
      支持以下 TYPE：
        INT / REAL / STR / BOOL / DICT    ← 标量
        [INT] [STR] [REAL] [BOOL]         ← 同行多值列表
        []                                ← 内嵌表格（子区域）
        {}                                ← 内嵌 key-value 表格（子区域）
        {T}                               ← 内嵌转置表格（子区域转置后按 key-value 解析）
        {DICT}                            ← 内嵌 key-value dict（子区域）

  表格格式（Sheet 名以 T_ 开头）：
      首行 = 字段名，次行 = 类型，其余为数据行
      若字段名中有 ID 列，返回 Dict[str, record]；否则返回 List[record]

## 特殊约定

  - Sheet 名以 '#' 开头的跳过
  - Sheet 名以 'T_' 开头的整张 Sheet 作为表格解析（_load_table）
  - 其余 Sheet 按 key-value 格式解析（_load_dict）
  - key 名为 Nameless 时，其子内容直接合并（inline）到父节点，该 key 本身不出现在结果中
    - 子内容为 dict（{} {T} {DICT}）：各子 key 直接写入父 dict
    - 子内容为 list（[] 或 [T]）：不支持，将抛出 TypeError

## 对外接口

    excel2dict(path) -> Dict[str, Any]
        将整个 Excel 解析为 {sheet_name: parsed_data}

    excel2json(path, output=None, indent=2) -> str
        将整个 Excel 解析后序列化为 JSON 字符串，可选写入文件
"""

from __future__ import annotations

import json
from pathlib import Path
from typing import Any, Dict, List, Union

try:
    from openpyxl import load_workbook
    HAS_OPENPYXL = True
except ImportError:
    HAS_OPENPYXL = False

# ─────────────────────────────────────────────
#  类型常量（与 slot_simulator_1 保持一致）
# ─────────────────────────────────────────────
_INT          = "INT"
_FLOAT        = "REAL"
_STR          = "STR"
_BOOL         = "BOOL"
_DICT         = "DICT"
_LIST_PREFIX  = ".."       # ..INT  ..STR 等（兼容旧约定）
_TABLE_SIGN   = "T_"       # Sheet 名前缀，表示整张 Sheet 是表格
_SKIP         = "#"        # Sheet 名或字段名前缀，表示跳过
_TRANSPOSE    = "T"        # {T} 转置标记
_NAMELESS     = "Nameless" # 特殊 key：内容直接合并到父节点，自身不输出


# ═══════════════════════════════════════════════════════
#  内部解析函数
# ═══════════════════════════════════════════════════════

def _cell_val(tp: str, value: Any, location: str = "") -> Any:
    """
    将单元格原始值按 tp 类型转换。
    location 仅用于错误提示。
    """
    # ── 列表类型：[INT] / [STR] / [REAL] / [BOOL] 或旧版 ..INT 等 ──
    if tp.startswith("[") and tp.endswith("]"):
        sub_tp = tp[1:-1]
        if sub_tp == "":
            # [] 由 _load_dict 提前拦截处理内嵌表格，此处不应出现
            raise TypeError(f"[{location}] [] 类型需在 _load_dict 中处理，请检查调用路径")
        if value is None or value == "":
            return []
        s = str(value).strip()
        if s.startswith("["):
            try:
                data = json.loads(s)
            except ValueError:
                raise ValueError(
                    f"[{location}] JSON 列表格式有误，请检查括号和引号: {s!r}"
                )
            return [_scalar(sub_tp, item, location) for item in data]
        parts = s.split(",")
        return [_scalar(sub_tp, p.strip().strip('"').strip("'"), location)
                for p in parts]

    # ── 旧版列表前缀：..INT / ..STR 等 ──
    if tp.startswith(_LIST_PREFIX):
        sub_tp = tp[len(_LIST_PREFIX):]
        if value is None or value == "":
            return []
        s = str(value).strip()
        if s.startswith("["):
            try:
                data = json.loads(s)
            except ValueError:
                raise ValueError(
                    f"[{location}] JSON 列表格式有误: {s!r}"
                )
            return [_scalar(sub_tp, item, location) for item in data]
        parts = s.split(",")
        return [_scalar(sub_tp, p.strip().strip('"').strip("'"), location)
                for p in parts]

    return _scalar(tp, value, location)


def _scalar(tp: str, value: Any, location: str = "") -> Any:
    """单个标量转换，None/空字符串返回类型默认值。"""
    if value is None or value == "":
        defaults = {_INT: 0, _STR: "", _FLOAT: 0.0, _BOOL: False, _DICT: {}}
        if tp in defaults:
            return defaults[tp]
        raise TypeError(f"[{location}] 未知类型 {tp!r}")

    if tp == _INT:
        try:
            return int(value)
        except (ValueError, TypeError):
            raise ValueError(f"[{location}] 无法转为 INT: {value!r}")
    if tp == _STR:
        return str(value)
    if tp == _FLOAT:
        try:
            return float(value)
        except (ValueError, TypeError):
            raise ValueError(f"[{location}] 无法转为 REAL: {value!r}")
    if tp == _BOOL:
        if value in (True, 1, "1", "TRUE", "true"):
            return True
        if value in (False, 0, "0", "FALSE", "false"):
            return False
        raise ValueError(f"[{location}] BOOL 只接受 0/1/TRUE/FALSE，got: {value!r}")
    if tp == _DICT:
        if isinstance(value, dict):
            return value
        try:
            return json.loads(str(value))
        except ValueError:
            raise ValueError(f"[{location}] DICT JSON 格式有误: {value!r}")
    raise TypeError(f"[{location}] 未知类型 {tp!r}")


def _count_col(rng: List[List], row: int = 0, col: int = 0,
               reverse: bool = False) -> int:
    """从 (row, col) 开始，计算连续非空（reverse=False）或非空（reverse=True）单元格数量。"""
    count = 0
    for i in range(col, len(rng[row])):
        is_empty = (rng[row][i] == "" or rng[row][i] is None)
        if reverse:
            if not is_empty:
                break
        else:
            if is_empty:
                break
        count += 1
    return count


def _count_row(rng: List[List], row: int = 0, col: int = 0,
               reverse: bool = False) -> int:
    """从 (row, col) 开始，计算连续非空（reverse=False）或非空（reverse=True）行数。"""
    count = 0
    for i in range(row, len(rng)):
        is_empty = (rng[i][col] == "" or rng[i][col] is None)
        if reverse:
            if not is_empty:
                break
        else:
            if is_empty:
                break
        count += 1
    return count


def _transpose(rng: List[List]) -> List[List]:
    """转置二维列表。"""
    if not rng or not rng[0]:
        return rng
    rows, cols = len(rng), max(len(r) for r in rng)
    return [
        [rng[r][c] if c < len(rng[r]) else None for r in range(rows)]
        for c in range(cols)
    ]


def _row_is_blank(rng: List[List], row_idx: int) -> bool:
    """判断整行是否全空（None / '' / 纯空格）。"""
    if row_idx >= len(rng):
        return True
    return all(
        v is None or (isinstance(v, str) and v.strip() == "")
        for v in rng[row_idx]
    )


def _is_key_row(row: list) -> bool:
    """
    判断一行是否是 key-value 行（而非表格数据行/类型行）。
    key-value 行特征：col 0 是普通名字，col 1 是类型标注。
    """
    if not row or len(row) < 2:
        return False
    c0 = row[0]
    c1 = row[1]
    # col 0 必须非空
    if c0 is None or (isinstance(c0, str) and c0.strip() == ""):
        return False
    # col 0 看起来像类型标注（如 [STR]、{T}）→ 这是表格的类型行，不是 key 行
    key_str = str(c0).strip()
    if (key_str.startswith("[") and key_str.endswith("]")) or \
       (key_str.startswith("{") and key_str.endswith("}")):
        return False
    if key_str.startswith(_LIST_PREFIX):
        return False
    # col 1 必须看起来像类型标注
    if c1 is None or (isinstance(c1, str) and c1.strip() == ""):
        return False
    tp = str(c1).strip()
    _builtin = {_INT, _FLOAT, _STR, _BOOL, _DICT, "[]"}
    if tp in _builtin:
        return True
    if (tp.startswith("[") and tp.endswith("]")) or \
       (tp.startswith("{") and tp.endswith("}")):
        return True
    if tp.startswith(_LIST_PREFIX):
        return True
    return False


def _get_sub_rng(rng: List[List], start_row: int):
    """
    从 start_row 开始提取子表格区域。

    算法（改进版）：
      1. 从 start_row+1 开始正向扫描，遇到"新 key 行"
         （_is_key_row 返回 True）或 sheet 末尾则停止。
      2. 提取 [start_row:scan_end][2:] 作为原始子区域。
      3. 从尾部删除全空行（避免空行被 _load_table 解析为空记录）。
      4. 计算最大有内容列数，裁剪右侧全空列。

    参数：
        rng        完整 Sheet 的二维列表
        start_row  子表格起始行（即 KEY 行的行号）

    返回：
        (sub_rng, row_count)
          sub_rng     裁剪后的二维列表（从原表 col 2 开始）
          row_count   从 start_row 到 scan_end 的原始行数（供外层 _load_dict
                      跳过已消费的行，不受尾部裁剪影响）
    """
    # ── 1. 正向扫描：找到子表格结束位置 ──
    scan_end = len(rng)  # 默认到 sheet 末尾
    for i in range(start_row + 1, len(rng)):
        if _is_key_row(rng[i]):
            scan_end = i
            break

    raw_row_count = scan_end - start_row  # 含 key 行的总行数

    # ── 2. 提取子区域（跳过 col 0,1） ──
    raw = [rng[i][2:] for i in range(start_row, scan_end)]

    # ── 3. 裁尾部全空行（保留至少表头行） ──
    while len(raw) > 1:
        # 检查 raw 的最后一行的原始行（即 scan_end-1 行）是否全空
        last_orig_idx = start_row + len(raw) - 1
        if _row_is_blank(rng, last_orig_idx):
            raw.pop()
        else:
            break

    if not raw:
        return [], raw_row_count

    # ── 4. 计算水平有效宽度 & 裁剪右侧全空列 ──
    max_col = 0
    for row_data in raw:
        for ci in range(len(row_data) - 1, -1, -1):
            v = row_data[ci]
            if v is not None and (not isinstance(v, str) or v.strip() != ""):
                col_end = ci + 1
                if col_end > max_col:
                    max_col = col_end
                break

    if max_col == 0:
        return [], raw_row_count

    trimmed = [r[:max_col] for r in raw]

    return trimmed, raw_row_count


def _load_table(rng: List[List],
                location: str = "") -> Union[List[Dict], Dict[str, Dict]]:
    """
    解析表格区域（首行字段名，次行类型，其余为数据行）。
    若字段中有 'ID' 列，返回 Dict[id_str, record]；否则返回 List[record]。
    """
    col_num = _count_col(rng)
    row_num = _count_row(rng)
    if col_num == 0 or row_num < 2:
        return []

    fields = rng[0][:col_num]
    types  = rng[1][:col_num]
    use_id = "ID" in fields
    result: Any = {} if use_id else []

    for r_idx in range(2, row_num):
        row = rng[r_idx]
        record: Dict[str, Any] = {}
        uid = None
        for c_idx, field in enumerate(fields):
            raw = row[c_idx] if c_idx < len(row) else None
            loc = f"{location}[R{r_idx + 1},C{c_idx + 1}]"
            if field == "ID":
                uid = str(_scalar(types[c_idx], raw, loc))
            elif field and field != _SKIP:
                record[field] = _cell_val(types[c_idx], raw, loc)
        if use_id:
            if uid in result:
                raise IndexError(f"[{location}] ID {uid!r} 重复")
            result[uid] = record
        else:
            result.append(record)

    return result


def _merge_nameless(result: Dict[str, Any], value: Any, loc: str) -> None:
    """
    将 Nameless key 的值合并到父 dict。
    - value 为 dict：各子 key 直接写入 result（key 冲突时后者覆盖前者）
    - value 为其他：抛出 TypeError
    """
    if not isinstance(value, dict):
        raise TypeError(
            f"[{loc}] Nameless key 的值必须是 dict 类型（{{}} / {{T}} / {{DICT}}），"
            f"实际得到 {type(value).__name__!r}"
        )
    result.update(value)


def _load_dict(rng: List[List],
               location: str = "") -> Dict[str, Any]:
    """
    解析 key-value 区域。
    每行格式：col0=key, col1=type, col2+=value

    特殊规则：
      - key 为 Nameless 时，将解析出的子 dict 内容直接合并到当前层级，
        Nameless 本身不出现在结果中。
    """
    result: Dict[str, Any] = {}
    row = 0
    while row < len(rng):
        key = rng[row][0]
        if key is not None and key != "":
            tp_raw = rng[row][1] if len(rng[row]) > 1 else None
            if tp_raw is None:
                row += 1
                continue
            tp  = str(tp_raw).strip()
            loc = f"{location}[key={key!r},R{row + 1}]"

            # ── 内嵌表格：[] ──
            if tp == "[]":
                sub_rng, row_count = _get_sub_rng(rng, row)
                val = _load_table(sub_rng, loc)
                row += row_count - 1
                if key == _NAMELESS:
                    _merge_nameless(result, val, loc)
                else:
                    result[key] = val

            # ── 同行多值列表：[INT] [STR] 等 ──
            elif tp.startswith("[") and tp.endswith("]"):
                sub_tp = tp[1:-1]
                values = rng[row][2: 2 + _count_col(rng, row, 2)]
                val = [_cell_val(sub_tp, v, loc) for v in values]
                if key == _NAMELESS:
                    _merge_nameless(result, val, loc)
                else:
                    result[key] = val

            # ── 嵌套 dict/表格：{} {T} {DICT} ──
            elif tp.startswith("{") and tp.endswith("}"):
                sub_rng, row_count = _get_sub_rng(rng, row)
                inner = tp[1:-1].strip()
                row += row_count - 1
                if inner == "":
                    val = _load_table(sub_rng, loc)
                elif inner == _TRANSPOSE:
                    val = _load_dict(_transpose(sub_rng), loc)
                elif inner == _DICT:
                    val = _load_dict(sub_rng, loc)
                else:
                    raise TypeError(
                        f"[{loc}] 未知嵌套类型 {tp!r}，支持: {{}} {{T}} {{DICT}}"
                    )
                if key == _NAMELESS:
                    _merge_nameless(result, val, loc)
                else:
                    result[key] = val

            # ── 旧版列表前缀 ..TYPE ──
            elif tp.startswith(_LIST_PREFIX):
                values = rng[row][2: 2 + _count_col(rng, row, 2)]
                val = [_cell_val(tp, v, loc) for v in values]
                if key == _NAMELESS:
                    _merge_nameless(result, val, loc)
                else:
                    result[key] = val

            # ── 标量 ──
            else:
                raw_val = rng[row][2] if len(rng[row]) > 2 else None
                val = _cell_val(tp, raw_val, loc)
                if key == _NAMELESS:
                    _merge_nameless(result, val, loc)
                else:
                    result[key] = val

        row += 1
    return result


def _read_sheet(ws) -> List[List]:
    """将 openpyxl worksheet 转为二维列表，datetime 转字符串。"""
    from datetime import datetime, time as dt_time
    rows = []
    for row in ws.rows:
        row_data = []
        for cell in row:
            v = cell.value
            if isinstance(v, (datetime, dt_time)):
                v = str(v)
            row_data.append(v)
        rows.append(row_data)
    return rows


# ═══════════════════════════════════════════════════════
#  对外接口
# ═══════════════════════════════════════════════════════

def excel2dict(path: str | Path) -> Dict[str, Any]:
    """
    将整个 Excel 文件解析为嵌套 dict。

    返回结构：
        {
            "SheetName1": { ... },   # key-value 格式的 Sheet → dict
            "SheetName2": [ ... ],   # T_ 前缀的表格 Sheet → list 或 dict
            ...
        }

    约定：
        - Sheet 名以 '#' 开头的跳过
        - Sheet 名以 'T_' 开头的整张 Sheet 按表格格式解析（_load_table）
        - 其余 Sheet 按 key-value 格式解析（_load_dict）
    """
    if not HAS_OPENPYXL:
        raise ImportError("请先安装 openpyxl: pip install openpyxl")

    wb = load_workbook(str(path), data_only=True)
    result: Dict[str, Any] = {}
    for ws in wb.worksheets:
        name = ws.title
        if name.startswith(_SKIP):
            continue
        rng = _read_sheet(ws)
        loc = f"Sheet({name!r})"
        if name.startswith(_TABLE_SIGN):
            result[name[len(_TABLE_SIGN):]] = _load_table(rng, loc)
        else:
            result[name] = _load_dict(rng, loc)
    return result


def excel2json(
    path: str | Path,
    output: str | Path | None = None,
    indent: int = 2,
) -> str:
    """
    将整个 Excel 文件解析为 JSON 字符串，可选写入文件。

    参数：
        path    Excel 文件路径（.xlsx）
        output  若提供，将 JSON 写入此路径
        indent  JSON 缩进空格数（默认 2）

    返回：
        JSON 字符串
    """
    data = excel2dict(path)
    text = json.dumps(data, ensure_ascii=False, indent=indent)
    if output:
        Path(output).write_text(text, encoding="utf-8")
    return text


# ═══════════════════════════════════════════════════════
#  CLI
# ═══════════════════════════════════════════════════════

if __name__ == "__main__":
    import sys

    if len(sys.argv) < 2:
        print("用法: python excel2json.py <excel_path> [output.json]")
        sys.exit(1)

    excel_path = sys.argv[1]
    output_path = sys.argv[2] if len(sys.argv) > 2 else None

    json_str = excel2json(excel_path, output=output_path)

    if output_path:
        print(f"已写入: {output_path}")
    else:
        print(json_str)
