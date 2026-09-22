#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""腾讯文档在线表格导出器（docs.qq.com/sheet）

把「通过分享链接匿名可读」的公开表格导出为 CSV。
不绕过文档访问权限——仅使用页面自身预加载的公开接口，不需要登录、不需要导出权限，
也不依赖浏览器（表格是 canvas 渲染，无障碍树里读不到内容）。

用法:
    python qq_sheet_export.py <url|docid> [tabid] [选项]

    <url>    完整的分享链接，如 https://docs.qq.com/sheet/XXXX?tab=yyyy
    <docid>  文档 ID（链接里 /sheet/ 后面那段）
    tabid    工作表 ID；省略则自动取第一个工作表

选项:
    -o, --out PATH      输出 CSV 路径（默认 <名称>.csv）
    --list-tabs         只列出文档里所有工作表（id + 名称）后退出
    --xlsx              额外用 format_excel 生成排版好的 xlsx
    --timeout N         单次请求超时秒数（默认 40）

只依赖 Python 标准库。
"""
import argparse
import base64
import csv
import json
import os
import re
import struct
import sys
import urllib.parse
import urllib.request
import zlib

API = 'https://docs.qq.com/dop-api/opendoc'
UA = ('Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 '
      '(KHTML, like Gecko) Chrome/140.0.0.0 Safari/537.36')
CHUNK_ROWS = 512          # 单次请求最多拉取的行数（分页粒度）
POOL_MIN_LEN = 10000      # 数据块里"主 f5"的最小长度


class ExportError(RuntimeError):
    """带可读说明的失败原因。"""


# --------------------------------------------------------------------------
# protobuf 基础（无 schema，按 wire format 扫描）
# --------------------------------------------------------------------------
def pb_varint(buf, i):
    shift = result = 0
    while i < len(buf):
        byte = buf[i]
        i += 1
        result |= (byte & 0x7F) << shift
        if not byte & 0x80:
            return result, i
        shift += 7
        if shift > 63:
            return None, i
    return None, i


def pb_fields(buf):
    """把一段字节切成 [(field_number, kind, value)]，kind ∈ v/l/f32/f64。"""
    out, i = [], 0
    while i < len(buf):
        key, i = pb_varint(buf, i)
        if key is None:
            return out
        field, wire = key >> 3, key & 7
        if field == 0 or wire not in (0, 1, 2, 5):
            return out
        if wire == 0:
            value, i = pb_varint(buf, i)
            out.append((field, 'v', value))
        elif wire == 2:
            length, i = pb_varint(buf, i)
            if length is None or i + length > len(buf):
                return out
            out.append((field, 'l', buf[i:i + length]))
            i += length
        elif wire == 5:
            out.append((field, 'f32', struct.unpack('<f', buf[i:i + 4])[0]))
            i += 4
        else:
            out.append((field, 'f64', struct.unpack('<d', buf[i:i + 8])[0]))
            i += 8
    return out


def pb_msg(buf, field, minlen=0):
    for num, kind, value in pb_fields(buf):
        if num == field and kind == 'l' and len(value) >= minlen:
            return value
    return None


def _utf8(b):
    try:
        return b.decode('utf-8')
    except UnicodeDecodeError:
        return ''


def pool_text(inner):
    """纯文本池条目：f1 → f1 = utf8"""
    for field, kind, value in pb_fields(inner):
        if field == 1 and kind == 'l':
            return _utf8(value)
    return ''


def pool_rich(inner):
    """富文本池条目：重复的 f3(文本段) → f3(文本容器) → f1 = 正文。

    样式（字体名/颜色/字号）在同级的 f1 里，必须按结构跳过——
    千万不要把整段可读字符串拼起来再清洗，那会把正文首字符一起删掉。
    """
    parts = []
    for field, kind, value in pb_fields(inner):
        if field != 3 or kind != 'l':
            continue
        for f2, k2, v2 in pb_fields(value):
            if f2 != 3 or k2 != 'l':
                continue
            for f3, k3, v3 in pb_fields(v2):
                if f3 == 1 and k3 == 'l':
                    parts.append(_utf8(v3))
    return ''.join(parts)


# --------------------------------------------------------------------------
# 取数
# --------------------------------------------------------------------------
def fetch(docid, tabid, start_row=0, end_row=0, end_col=63, timeout=40):
    query = urllib.parse.urlencode({
        'tab': tabid or '', 'id': docid, 'u': '', 'noEscape': 1,
        'enableSmartsheetSplit': 1, 'needSheetState': 1, 'sliceStates': 1,
        'block_start_col': 0, 'block_start_row': start_row,
        'block_end_col': end_col, 'block_end_row': end_row,
        'startrow': 0, 'endrow': 1000, 'normal': 1, 'outformat': 1,
        'wb': 1, 'nowb': 0, 'callback': 'clientVarsCallback', 'xsrf': '',
    })
    req = urllib.request.Request(
        API + '?' + query,
        headers={'User-Agent': UA,
                 'Referer': 'https://docs.qq.com/sheet/' + docid})
    try:
        body = urllib.request.urlopen(req, timeout=timeout).read()
    except Exception as exc:                      # noqa: BLE001
        raise ExportError('请求失败（网络/超时/被拦截）: %s' % exc)
    text = body.decode('utf-8', 'replace')
    match = re.match(r'^[^(]*\((.*)\)\s*$', text, re.S)
    if not match:
        raise ExportError(
            '返回的不是预期 JSONP——通常意味着文档不可匿名访问、ID 不对，或接口已变更。'
            '返回前 200 字符: %r' % text[:200])
    try:
        payload = json.loads(match.group(1))
    except ValueError as exc:
        raise ExportError('JSONP 里的 JSON 解析失败: %s' % exc)

    node = payload.get('clientVars', {}).get('collab_client_vars', {})
    text_list = node.get('initialAttributedText', {}).get('text') or []
    if not text_list or not isinstance(text_list[0], dict):
        raise ExportError('返回里找不到 initialAttributedText.text[0]——接口可能已变更')
    return text_list[0]


def list_sheets(docid, timeout=40):
    """枚举文档里所有工作表 → [(tabid, 名称)]"""
    info = fetch(docid, None, 0, 0, timeout=timeout)
    workbook = info.get('workbook')
    if not workbook:
        raise ExportError('返回里没有 workbook 字段，无法枚举工作表')
    raw = zlib.decompress(base64.b64decode(workbook))
    top = pb_msg(raw, 1)
    if top is None:
        raise ExportError('workbook 结构不符预期（接口可能已变更）')
    sheets = []
    for field, kind, value in pb_fields(top):
        if field != 5 or kind != 'l':
            continue
        tab_id = name = None
        for f2, k2, v2 in pb_fields(value):
            if f2 != 2 or k2 != 'l':
                continue
            for f3, k3, v3 in pb_fields(v2):
                if f3 == 3 and k3 == 'l':
                    tab_id = _utf8(pb_msg(v3, 1) or b'') or None
                elif f3 == 5 and k3 == 'l':
                    name = _utf8(pb_msg(v3, 1) or b'') or None
        if tab_id:
            sheets.append((tab_id, name or tab_id))
    if not sheets:
        raise ExportError('workbook 里没解析到工作表（接口可能已变更）')
    return sheets


def decode_block(related_b64):
    """解开一个数据块 → [(row, col, tp, idx)] + 三个值池"""
    raw = zlib.decompress(base64.b64decode(related_b64))
    top = pb_msg(raw, 1)
    if top is None:
        raise ExportError('数据块结构不符预期')
    big = None
    for field, kind, value in pb_fields(top):
        if field == 5 and kind == 'l' and len(value) > POOL_MIN_LEN:
            big = value
    if big is None:
        raise ExportError('数据块里找不到主数据段（空表，或接口已变更）')
    node = pb_msg(big, 19)
    if node is None:
        raise ExportError('数据块里找不到 f19 段（接口可能已变更）')

    fields = pb_fields(node)
    pool = None
    for field, kind, value in fields:
        if field == 5 and kind == 'l':
            pool = value
            break
    if pool is None:
        raise ExportError('数据块里找不到值池（接口可能已变更）')

    texts, rich, nums = [], [], []
    i = 0
    while i < len(pool):
        key, i = pb_varint(pool, i)
        if key is None:
            break
        field, wire = key >> 3, key & 7
        if field == 0 or wire not in (0, 1, 2, 5):
            break
        if wire == 2:
            length, i = pb_varint(pool, i)
            if length is None or i + length > len(pool):
                break
            inner = pool[i:i + length]
            i += length
            if field == 1:
                texts.append(pool_text(inner))
            elif field == 2:
                rich.append(pool_rich(inner))
            elif field == 3:
                sub = pb_fields(inner)
                nums.append(sub[0][2] if sub else None)
        elif wire == 0:
            _, i = pb_varint(pool, i)
        elif wire == 5:
            i += 4
        else:
            i += 8

    cells = []
    for _, kind, value in fields:
        if kind != 'l':
            continue
    for cell in [v for f, k, v in fields if f == 6 and k == 'l']:
        row = col = 0
        tp = idx = None
        for f2, k2, v2 in pb_fields(cell):
            if f2 == 1 and k2 == 'v':
                row = v2
            elif f2 == 2 and k2 == 'v':
                col = v2
            elif f2 == 3 and k2 == 'l':
                for f3, k3, v3 in pb_fields(v2):
                    if f3 == 1 and k3 == 'v':
                        tp = v3
                    elif f3 == 2 and k3 == 'l':
                        sub = pb_fields(v3)
                        idx = 0 if not sub else sub[0][2]
        cells.append((row, col, tp, idx))
    return cells, texts, rich, nums


def cell_value(tp, idx, texts, rich, nums):
    """按值类型还原单元格文本。返回 None 表示该格无内容。"""
    if idx is None:
        return None
    if tp == 4:                                   # 纯文本
        return texts[idx] if 0 <= idx < len(texts) else None
    if tp == 6:                                   # 富文本
        return rich[idx] if 0 <= idx < len(rich) else None
    if tp == 2:                                   # 数字
        if idx < 129:
            value = idx
        else:
            pos = idx - 129
            if pos >= len(nums) or nums[pos] is None:
                return None
            value = nums[pos]
        if isinstance(value, float) and value == int(value):
            return str(int(value))
        return str(value)
    if tp == 5:                                   # 自动编号 / 公式（如「序号」列）
        return str(idx + 1)
    return None


def export(docid, tabid, timeout=40, verbose=True):
    """拉取整张表 → (grid, meta)。分页拉取，覆盖 max_row 全部行。"""
    probe = fetch(docid, tabid, 0, 0, timeout=timeout)
    max_row = probe.get('max_row')
    max_col = probe.get('max_col')
    if not isinstance(max_row, int) or max_row <= 0:
        raise ExportError('没拿到有效的 max_row（表可能为空，或接口已变更）')
    if verbose:
        print('  表尺寸: %d 行 × %s 列' % (max_row, max_col), file=sys.stderr)

    grid = {}
    done = set()
    start = 0
    while start < max_row:
        end = min(start + CHUNK_ROWS - 1, max_row - 1)
        if verbose and start:
            print('  续取第 %d-%d 行…' % (start, end), file=sys.stderr)
        info = fetch(docid, tabid, start, end, max_col + 1, timeout=timeout)
        blocks = info.get('block_datas') or []
        if not blocks:
            raise ExportError('第 %d-%d 行没有返回任何数据块' % (start, end))
        for block in blocks:
            related = block.get('related_sheet')
            if not related:
                continue
            cells, texts, rich, nums = decode_block(related)
            for row, col, tp, idx in cells:
                key = (row, col)
                if key in done:
                    continue
                value = cell_value(tp, idx, texts, rich, nums)
                if value is not None:
                    grid[key] = value
                done.add(key)
        start = end + 1

    if not grid:
        raise ExportError('没解析到任何单元格')
    nrow = max(r for r, _ in grid) + 1
    ncol = max(c for _, c in grid) + 1
    out = [[''] * ncol for _ in range(nrow)]
    for (row, col), value in grid.items():
        out[row][col] = value
    return out, {'max_row': max_row, 'max_col': max_col, 'cells': len(grid)}


# --------------------------------------------------------------------------
# CLI
# --------------------------------------------------------------------------
def parse_target(text):
    """把完整链接或裸 ID 解析成 (docid, tabid)"""
    if 'docs.qq.com' in text:
        match = re.search(r'/sheet/([A-Za-z0-9]+)', text)
        if not match:
            raise ExportError('只支持 docs.qq.com/sheet/ 表格链接: %s' % text)
        query = urllib.parse.parse_qs(urllib.parse.urlparse(text).query)
        return match.group(1), (query.get('tab') or [None])[0]
    return text, None


def main(argv=None):
    ap = argparse.ArgumentParser(
        description='导出腾讯文档公开表格为 CSV（仅限匿名可读的分享链接）')
    ap.add_argument('target', help='分享链接或文档 ID')
    ap.add_argument('tabid', nargs='?', help='工作表 ID（默认第一个）')
    ap.add_argument('-o', '--out', help='输出 CSV 路径')
    ap.add_argument('--list-tabs', action='store_true', help='只列出所有工作表')
    ap.add_argument('--xlsx', action='store_true', help='同时生成排版好的 xlsx')
    ap.add_argument('--timeout', type=int, default=40)
    args = ap.parse_args(argv)

    try:
        docid, tabid = parse_target(args.target)
        tabid = args.tabid or tabid

        if args.list_tabs:
            print('文档 %s 的工作表：' % docid)
            for tid, name in list_sheets(docid, timeout=args.timeout):
                print('  %-10s %s' % (tid, name))
            return 0

        name = tabid
        if not tabid:
            sheets = list_sheets(docid, timeout=args.timeout)
            tabid, name = sheets[0]
            print('未指定 tab，自动使用: %s (%s)' % (name, tabid), file=sys.stderr)

        grid, meta = export(docid, tabid, timeout=args.timeout)
        out_path = args.out or '%s.csv' % (name or tabid)
        with open(out_path, 'w', newline='', encoding='utf-8-sig') as fh:
            csv.writer(fh).writerows(grid)
        print('%s  ←  %d 行 × %d 列（%d 个非空单元格）'
              % (out_path, len(grid), len(grid[0]), meta['cells']))

        if args.xlsx:
            try:
                from format_excel import csv_to_xlsx
            except ModuleNotFoundError as exc:
                if exc.name == 'format_excel':
                    print('跳过 --xlsx：同目录下找不到 format_excel.py', file=sys.stderr)
                elif exc.name == 'openpyxl':
                    print('跳过 --xlsx：缺少可选依赖 openpyxl；请运行 pip install openpyxl',
                          file=sys.stderr)
                else:
                    raise
            else:
                xlsx_path = os.path.splitext(out_path)[0] + '.xlsx'
                csv_to_xlsx(out_path, xlsx_path)
                print(xlsx_path)
        return 0
    except ExportError as exc:
        print('错误: %s' % exc, file=sys.stderr)
        return 1


if __name__ == '__main__':
    sys.exit(main())
