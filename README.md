# 腾讯文档表格导出器（docs.qq.com/sheet）

把**通过分享链接匿名可读**的腾讯文档表格导出为 CSV / Excel。

页面上的「导出为 Excel」经常被禁用，表格又是 canvas 渲染（浏览器自动化抓不到、无障碍树里没有单元格）。
这个工具直接使用页面**自身预加载的公开接口**取出 protobuf 数据再解码——不需要登录、不需要导出权限、不需要浏览器，**只依赖 Python 标准库**。

---

## ⚠️ 使用边界（请先读这一节）

- 本工具**只能读取「匿名可读」的公开分享文档** —— 也就是你不登录、打开链接就能看到内容的那种。
- 它**不绕过任何访问权限**：需要登录、需要授权、需要密码的文档一律读不到，本项目也无意支持这类用法。
- 它调用的是**前端内部接口**（`/dop-api/opendoc`），**不是**腾讯文档的公开 API。**这可能与腾讯文档的服务条款相冲突**，是否使用请自行判断并承担相应责任。
- 接口格式是**逆向**得来的，**随时可能变化**。今天能用不代表明天能用（见下文「自检」）。
- 请尊重文档作者/分享者的意愿，并遵守相关法律法规。

---

## 快速开始

需要 Python 3.8+，**无第三方依赖**。

```bash
# 直接给链接（自动解析文档 ID 和 tab）
python qq_sheet_export.py "https://docs.qq.com/sheet/YOUR_DOC_ID?tab=YOUR_TAB_ID"

# 也可以只给文档 ID，再用 --list-tabs 看有哪些工作表
python qq_sheet_export.py --list-tabs YOUR_DOC_ID
python qq_sheet_export.py YOUR_DOC_ID YOUR_TAB_ID

# 指定输出路径 / 顺带生成排版好的 xlsx
python qq_sheet_export.py "https://docs.qq.com/sheet/YOUR_DOC_ID?tab=YOUR_TAB_ID" -o data.csv --xlsx
```

输出为 **UTF-8-SIG** 的 CSV（Excel 双击不乱码）。

排版成 Excel（表头样式、斑马纹、边框、列宽自适应、冻结首行、自动筛选）需要 `openpyxl`：

```bash
pip install openpyxl
python format_excel.py data.csv              # → data.xlsx
python format_excel.py data.csv pretty.xlsx
```

---

## 它做了什么

```
docs.qq.com/dop-api/opendoc      ← 页面自身预加载的数据接口（JSONP）
        ↓  clientVars.collab_client_vars.initialAttributedText.text[0]
        ↓  block_datas[n].related_sheet
   base64 → zlib → protobuf
        ↓
   f19.f6[]  = 单元格（行号 / 列号 / 值类型 / 值索引）
   f19.f5    = 值池（文本池 / 富文本池 / 数字池）
```

细节（含踩过的坑）都写在 **[docs/schema.md](docs/schema.md)** —— 如果你需要自己修这个工具，那份文档比代码更重要。

几个关键点：

- **空单元格不落盘**，只能靠行列号定位，不能按出现顺序硬拼；
- **合并单元格只记左上角**（序号/学校等），导出后通常要向下填充；
- **富文本必须按结构取正文**（`段 → 文本容器 → f1`），样式（字体/颜色/字号）在旁路字段里。
  千万不要把所有可读字符串拼起来再清洗 —— 那会把正文首字符一起删掉（`26科目` 变成 `6科目`）；
- **数字池有两种形态**：索引 < 129 时索引即值，≥ 129 时指向数字池；
- **`tp=5` 是自动编号/公式**（常见于「序号」列），漏掉它会让整列空白。

---

## 已知限制

| 限制 | 说明 |
|---|---|
| 仅限匿名可读文档 | 需要登录/授权的文档读不到 |
| 仅支持**表格**（`/sheet/`） | `/doc/`、`/slide/`、`/pdf/`、`/form/` 是另一套结构，不支持 |
| 强依赖逆向结构 | 腾讯改版即失效，无解 |
| `tp=5` 的取值规则 | 「索引 + 1」是**归纳**出来的（不是官方定义），换文档时建议先核对第一条数据的序号 |
| 大表 | 已按 `max_row` 分页拉取；实测覆盖到千行级，更极端的表格若漏行，请提 issue 附上文档 ID |
| CSV 不做加工 | 保留原始值（含行内换行、尾随空格）。清理与排版交给 `format_excel.py` |

---

## 自检：怀疑接口变了，先看这三处

1. **表头有没有** `SimHei` / `imHei` / `FFFFFFFF` 之类残留 → 富文本解析错了
2. **表头数字有没有被吃掉**（`26科目` 变成 `6科目`）→ 用了「展平 + 字符串清洗」的错误做法
3. **序号列是不是整列空白** → 漏了 `tp=5`

再不行就跑 `python qq_sheet_export.py --list-tabs <docid>`：如果这一步就报错，说明接口层变了（错误信息里会给出返回内容的前 200 字符，便于比对）。

---

## 与同类项目的区别

GitHub 上已有一些腾讯文档解析项目，多数是**库 / CLI / 浏览器自动化**（需要 cookie 或 headless 浏览器）。
本项目的定位不同：**给 AI coding agent 用的一页式 playbook + 可直接运行的零依赖 fallback**，
重点是把它固化在 [SKILL.md](SKILL.md) 里——包括**为什么不能用字符串清洗**这类故障经验。
如果你要的是一个成熟的库，建议用那些项目；如果你要的是「agent 遇到 docs.qq.com 链接时不浪费时间」，这个更合适。

---

## License

MIT
