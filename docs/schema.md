# 数据结构笔记（逆向所得）

> 本文记录 `docs.qq.com/sheet` 前端预加载数据的内在结构。
> 它是**逆向**得出的，不是官方定义，腾讯改版后可能失效。
> 写在这里的目的：**万一结构变了，你能看懂代码在做什么，并自己修。**

## 1. 取数入口

页面 HTML 里 preload 了一个接口（在页面源码里 grep `dop-api/opendoc` 就能拿到完整参数）：

```
https://docs.qq.com/dop-api/opendoc?tab=<TABID>&id=<DOCID>&noEscape=1
  &enableSmartsheetSplit=1&needSheetState=1&sliceStates=1
  &block_start_col=0&block_start_row=0&block_end_col=63&block_end_row=1023
  &startrow=0&endrow=1000&normal=1&outformat=1&wb=1&nowb=0
  &callback=clientVarsCallback&xsrf=
```

| 参数 | 作用 |
|---|---|
| `tab` / `id` | 工作表 ID / 文档 ID |
| `block_start_row` / `block_end_row` | **行切片（闭区间，分页靠它）** |
| `block_start_col` / `block_end_col` | 列切片 |
| `outformat=1&wb=1&nowb=0` | workbook 形态返回 |
| `needSheetState=1&sliceStates=1` | 附带工作表状态 |
| `callback=clientVarsCallback` | JSONP 回调名 |

响应是 **JSONP**：`clientVarsCallback({...})`。

> `max_row` / `max_col` 描述的是**全表**，与 `block_*` 无关。
> 所以可以先用一次很小的请求探出尺寸，再分页拉全。

## 2. 响应骨架

```
clientVars
 └ collab_client_vars
    └ initialAttributedText.text[0]
       ├ max_row / max_col
       ├ workbook                  ← 工作表清单（base64 → zlib → protobuf）
       └ block_datas[]             ← 数据块（可有多个）
          ├ start_row_index / end_row_index / end_col_index
          └ related_sheet          ← 单元格数据（base64 → zlib → protobuf）
```

`related_sheet` 解码后：

```
top.f1 (msg)
 └ f5 (msg, 体积最大的那个)        ← 一个工作表的数据块
    └ f19 (msg)
       ├ f3  (msg)               元信息: f1=sheetId, f4=maxRow-1, f5=maxCol-1
       ├ f4  (msg)               列格式
       ├ f5  (msg)               值池  ← 见第 3 节
       └ f6  (msg) × N           单元格
             f1       = 行号 (0-based)
             f2       = 列号 (0-based)
             f3.f1    = 值类型 tp
             f3.f2.f1 = 值索引 idx
             f3.f4.f1 = 样式索引
```

**关键**：`f6` **只记录有内容的单元格**，空单元格根本不出现。
所以必须用 `f1`/`f2` 精确定位，**不能按出现顺序硬拼**。

## 3. 值池

`f19.f5` 内部是一串重复条目 `<tag> <len> <content>`，**tag 决定值类型**：

| tag | content 结构 | 单元格 tp | 取值 |
|---|---|---|---|
| `0a` (f1) | `f1 → f1 = utf8` | **4** | 直接就是文本 |
| `12` (f2) | 重复的 `f3`（文本段） | **6** | 见第 4 节 |
| `1a` (f3) | `f1 = fixed64`（double） | **2** | `idx < 129` → 值即 `idx`；否则 `池[idx-129]` |

还有一类单元格 **`tp = 5`**：自动编号 / 公式（常见于「序号」列）。
它的 `idx` 从 0 递增，**值 = idx + 1**。漏掉它会让整列空白。

> `129` 这个分界、以及 `tp=5` 的 `+1`，都是从样本**归纳**出来的经验规则，不是官方定义。

## 4. 富文本（tp=6）—— 最容易做错的地方

content 是**重复的 `f3` 文本段**，每段：

```
f3 (段)
  ├ f1 (样式)   f1.f1.f1     = 字体名   (e.g. 'SimHei')
  │             f1.f21.f5.f1 = 颜色     (e.g. 'FFFFFFFF')
  │             f1.f23.f1    = 字号
  ├ f3 (文本容器) → f1 = 正文      ← 只有这里是正文
  └ f5 = 3
```

正确取法：遍历段 → 取「文本容器 → f1」再拼接。**段之间直接拼，不要自己加分隔符**
（换行如果存在，本身就是某个段的文本，例如 `run = '\n\n'`）。

### 反例（真实踩坑，会损坏数据）

把 content 里所有可读字符串**展平**，再用正则删字体名/颜色码，
最后为了清掉样式字符而「删掉每段首字符」。后果：

- `SimHei` 删不干净 → 表头变成 `imHei\n6科目SimHei\nimHei\n（英/数/专）`
- 正文首字符被吃掉 → `26科目` 变 `6科目`、`26改考` 变 `7改考`

**这种损坏不可逆**——原始字符已经没有别的来源了。

**结论**：样式与正文在 protobuf 里是**不同字段**。用结构区分，绝不要用字符串猜测。

## 5. workbook（工作表清单）

`text[0].workbook` 解码后：

```
f1 (msg)
 └ f5[]                 ← 每个工作表一项
      f2.f3.f1 = sheetId   (e.g. 'BB08J2')
      f2.f5.f1 = 名称       (e.g. '985')
```

据此可自动枚举文档里所有 tab，无需人工翻页面。

## 6. 分页

实测（某 214 行文档）：

| 请求参数 | 返回 |
|---|---|
| `block_end_row=1023` | `max_row=214`，`end_row_index=213`（全表） |
| `block_end_row=6` | `max_row=214`，`end_row_index=6`（**只有 0–6 行**） |
| `block_start_row=7` | `start_row_index=7`，`end_row_index=213` |

结论：**`block_*` 是闭区间切片**，`max_row` 始终是全表维度。
因此「先探尺寸 → 分块请求 → 按 `(row, col)` 合并」即可覆盖整表。

## 7. 归因与免责

以上全部来自**对公开接口响应的观察**，未使用任何非公开文档或工具。
结构可能随时变化；本项目按 MIT 协议「as is」提供，不保证可用性。
