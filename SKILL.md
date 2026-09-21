---
name: tencent-docs-sheet
description: 导出 docs.qq.com 在线表格为 CSV/Excel——仅限「匿名可读」的公开分享文档，不绕过访问权限；含 protobuf 结构笔记与踩坑清单，附零依赖脚本
---

# 腾讯文档在线表格导出（docs.qq.com/sheet）

把 `https://docs.qq.com/sheet/<DOCID>?tab=<TABID>` 这类**公开分享链接**的表格数据导出为 CSV。
不需要登录、不需要「导出为 Excel」权限、不需要浏览器（表格是 canvas 渲染，无障碍树里只有工具栏，读不到单元格）。

**使用边界**：只适用于**不登录就能看到内容**的文档；**不绕过任何访问权限**。
调用的 `/dop-api/opendoc` 是前端内部接口（非公开 API），可能与服务条款冲突，接口也可能随时变更。

## 触发场景
- 用户给出 docs.qq.com/sheet 链接，要内容 / 要 CSV / 要表格数据
- 页面上的导出功能被禁用
- 需要批量导出同一文档的多个工作表

## 怎么做

**首选**：直接跑仓库里的脚本（零依赖，能自动解析链接、枚举 tab、分页拉全）：

```bash
python qq_sheet_export.py "<分享链接>"            # 自动取第一个工作表
python qq_sheet_export.py --list-tabs <docid>     # 先看有哪些工作表
python qq_sheet_export.py <docid> <tabid> -o out.csv --xlsx
```

需要自己实现时，按下面的结构来。

## 数据链路

```
GET https://docs.qq.com/dop-api/opendoc?tab=<TABID>&id=<DOCID>&noEscape=1
    &enableSmartsheetSplit=1&needSheetState=1&sliceStates=1
    &block_start_row=0&block_end_row=<max_row-1>&block_end_col=63
    &normal=1&outformat=1&wb=1&nowb=0&callback=clientVarsCallback&xsrf=
      ↓ JSONP: clientVarsCallback({...})
clientVars.collab_client_vars.initialAttributedText.text[0]
  ├ max_row / max_col              全表维度（与 block_* 无关）
  ├ workbook                       工作表清单（base64→zlib→protobuf）
  └ block_datas[].related_sheet    单元格数据（base64→zlib→protobuf）
```

`related_sheet` 结构：

```
top.f1 → f5(最大的那个) → f19
   ├ f3  元信息
   ├ f5  值池            ← 见下
   └ f6[] 单元格          f1=行号 f2=列号 f3.f1=类型tp f3.f2.f1=索引idx
```

**先探尺寸再分页**：`max_row` 是全表行数，且 `block_*` 是闭区间切片
（实测 `block_end_row=6` 就只返回 0–6 行）→ 按 512 行一档请求、按 `(row, col)` 合并即可覆盖整表。

## 值池与取值（核心）

池条目 = `<tag> <len> <content>`，按 tag 分三类：

| tag | 结构 | tp | 取值 |
|---|---|---|---|
| `0a` | `f1→f1=utf8` | **4** | 就是文本 |
| `12` | 重复 `f3` 文本段 | **6** | 见下（富文本） |
| `1a` | `f1=fixed64(double)` | **2** | `idx<129` 时 idx 即值，否则 `池[idx-129]` |

另有 **`tp=5`**：自动编号/公式（如「序号」列），idx 从 0 递增，**值 = idx+1**。

## 富文本（tp=6）必须按结构取，**不要做字符串清洗**

每段结构：`f3 { f1=样式(字体/颜色/字号), f3=文本容器 → f1=正文, f5=3 }`。
只取「文本容器 → f1」并按段拼接（**段之间直接拼，不要自己加分隔符**——换行本身就在某个段里）。

```python
def pool_rich(inner):
    parts = []
    for f, t, v in pb_fields(inner):          # 顶层：重复 f3 段
        if f != 3 or t != 'l': continue
        for a, b, c in pb_fields(v):          # 段内
            if a != 3 or b != 'l': continue   # 文本容器
            for x, y, z in pb_fields(c):
                if x == 1 and y == 'l':
                    parts.append(z.decode('utf-8'))
    return ''.join(parts)
```

**反例（真实损坏过数据）**：把所有可读字符串展平再正则清洗，并「删每段首字符」清样式 →
`SimHei` 残留成 `imHei`，正文首字符被吃掉（`26科目`→`6科目`、`26改考`→`7改考`），**不可逆**。

## 其他坑

1. **空单元格不落盘** —— 必须用行列号定位，不能按顺序硬拼。
2. **合并单元格只记左上角** —— 序号/学校/地区等只在块首有值；导出后一般要「向下填充」，且不同表列序可能不同，**按表头名定位要填充的列**，别写死列号。
3. **`tp=5` 最容易漏** —— 解码表只写 4/6/2 时，序号列会整列空白。
4. **浏览器那条路走不通** —— canvas 渲染，a11y 树里只有工具栏。
5. **开头常有非数据行** —— 图例说明行、广告行，按需剔除。
6. **CSV 用 `utf-8-sig`** —— Excel 打开才不乱码。
7. **写 xlsx 前清掉 XML 非法控制字符**（`[\x00-\x08\x0b\x0c\x0e-\x1f]`），否则 openpyxl 报错或文件损坏。

## 自检清单（导出后先看这三处）

1. **表头**有没有 `SimHei` / `imHei` / `FFFFFFFF` 残留 → 富文本解析错了
2. **表头数字**有没有被吃掉（`26科目` → `6科目`）→ 用了字符串清洗
3. **序号列**是不是整列空白 → 漏了 `tp=5`

再不行跑 `--list-tabs`：这一步就失败说明接口层变了（脚本会打印返回内容前 200 字符便于比对）。

## 不确定性说明

- 整套结构是**逆向**的，腾讯改版即失效；
- `tp=2` 的 `129` 分界、`tp=5` 的 `+1` 都是**归纳**出来的，不是官方定义 —— **换文档时先核对第一条数据的序号**；
- 已实测覆盖千行级表格；更极端的大表若漏行，请附文档 ID 反馈。
