# Excel 资产负债表到 Word 会计报告自动填报

这是一个可以本地运行、无需 API key 的 Python 项目。它从资产负债表提取金额，按显式配置匹配科目，在全部校验通过后填入 Word，并生成可追溯的核对日志。原始 Excel、Word 模板和配置文件保持原样。

当前交付采用可替换示例数据；没有使用你的真实模板。示例的公司、科目和金额仅用于验证流程，不代表完整企业会计报表或特定会计准则。更换真实数据时，需要同步调整映射、勾稽关系、公司、日期和单位。

## 先运行一次示例

需要 Python 3.10 或以上。核心运行只依赖 `openpyxl` 和 `lxml`，不要求安装 Microsoft Office、不要求联网调用 AI。首次安装依赖需要网络，依赖已安装后核心流程可以离线运行。

Windows 可双击项目根目录的 **run_demo.bat**。它会创建本项目专用的虚拟环境、安装依赖并运行示例。示例 Word 和日志保存到新的 `results/日期时间-随机编号/` 目录。如果 Windows 提示找不到 `py`，安装 Python 并勾选启动器，或按下面手动运行。若双击窗口关闭过快，可在终端运行该文件查看错误。

Windows PowerShell 手动运行：

```powershell
# 在解压后的项目根目录运行
py -3 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe demo.py
```

macOS / Linux：

```sh
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
.venv/bin/python demo.py
```

也可以安装为命令行工具：

```sh
python -m pip install -e .
accounting-report --help
```

以下命令使用当前环境的 `python`；Windows 虚拟环境可将它替换为 `.\.venv\Scripts\python.exe`。

```sh
python -m accounting_report run --excel examples/balance_sheet.xlsx --template examples/report_template.docx --config examples/mapping.json --output-dir results/my-first-run
```

**输出目录必须不存在**，同一命令重复执行时请更换目录名。这样可保留每次核对结果，不覆盖历史报告。`demo.py` 自动生成唯一目录名。

## 已包含的文件

- `accounting_report/`：四层核心源码、命令行入口和可选 AI 适配器。
- `examples/balance_sheet.xlsx`：左右两栏、两期余额、带合计公式及已保存计算结果的示例资产负债表。
- `examples/report_template.docx`：正文和表格占位符模板，含跨多个 Word run 的占位符。
- `examples/mapping.json`：32 个金额字段、16 项勾稽检查、元到万元的换算配置。
- `examples/create_templates.py`：可重新生成示例 Word 和映射；示例 Excel 已包含。
- `tests/test_project.py`：自动化测试；测试期间不调用真实 AI 服务。
- `demo_output/`：本次实际运行得到的 Word、JSON 和 CSV 核对日志。
- `requirements.txt`、`requirements-dev.txt`、`pyproject.toml`：运行和开发依赖。
- `docs/configuration.md`：配置字段及定位方法的详细说明。
- `VERIFICATION.md`：本次交付的验证结果与验证边界。

## 四层架构

1. **Excel 数据提取** `extract.py`：读取指定工作表及范围，支持同一张表的左右栏、不同期间列、科目名称或固定坐标。记录原始值、公式、来源单元格、科目和期间。
2. **规则映射与可选语义建议** `config.py`、`ai.py`：显式科目名和别名精确匹配。只去除空白、全半角差异及末尾冒号，不合并含义不同的科目。匹配不唯一会阻断填报。AI 和本地相似度建议是独立命令，不能进入自动填报链路。
3. **严格数据验证** `numbers.py`、`validate.py`：检查配置、来源标识、金额类型、空值、范围、科目唯一性、公式缓存及全部配置的勾稽关系。使用 `Decimal` 和四舍五入 `ROUND_HALF_UP`。勾稽比较使用换算前的原始金额，容差也使用 Excel 来源单位。
4. **Word 保留格式填报与核对** `word.py`、`pipeline.py`：直接编辑 DOCX 的 XML 文本节点，不重建全文段落或表格。支持正文、表格、嵌套表格、页眉和页脚的占位符；保留文本片段格式、段落属性、表格结构，以及未修改 ZIP 成员的原始字节。先写临时报告，验证保存后的文件后发布。

## 结果和日志怎么看

运行成功的目录包含：

- `report.docx`：填好的会计报告。
- `audit.json`：每个来源文件的 SHA-256、每个金额的来源坐标、原始值、公式、单位换算、舍入结果、全部勾稽检查、Word 替换记录、警告和错误。时间明确记录为 UTC。
- `audit.csv`：便于 Excel 查看和人工核对的逐字段明细。文本以 `=`、`+`、`-`、`@` 开头时会加单引号，避免打开 CSV 时把源公式当作可执行公式。

失败时目录中只留下核对日志，**不生成 `report.docx`**。如果某字段读取失败，JSON 的 `errors` 给出字段和原因，受影响检查标记为 `unavailable`，不会显示通过；CSV 只含成功提取字段。无法创建输出目录或目录已经存在时，错误显示在终端，程序不写入这个目录。

成功退出码为 `0`，验证或运行失败为 `2`。校验通过的预演状态是 `validated`，正常填报状态是 `completed`。

只验证、不生成 Word：

```sh
python -m accounting_report run --excel examples/balance_sheet.xlsx --template examples/report_template.docx --config examples/mapping.json --output-dir results/preflight --dry-run
```

预演会检查金额、勾稽关系和 Word 的全部目标，生成完整日志。它不能替代 Word 中的人工版面检查。

## 接入你的真实 Excel 和 Word

1. 把 Excel 另存为 `.xlsx`，Word 另存为 `.docx`。Excel 如含公式，先在 Excel 中重算并保存。可读取 `.xlsm` 的数值，但不运行宏、不修改源文件。
2. 复制 `examples/mapping.json` 为自己的配置，修改 `source.sheet`、`source.unit`、扫描行范围和期间列。双栏表使用多个 `blocks`；同一科目出现在多个分区时，用不同 block 限定，不跨分区猜测。
3. 配置 `metadata` 校验公司、金额单位、每个期间列的表头，确保没有拿错公司、报表、期初期末或金额单位。每次换报告期间应同时更新 Excel 表头、配置预期值和 Word 中的日期。
4. 修改每个字段的 `source.labels`，只填你已经确认含义等价的名称。需要固定坐标时，同时配置 `label_cell` 校验该坐标旁的科目，避免插行后读错金额。
5. 在 Word 需要填写的空行上，输入 `{{cash.end}}` 这样的占位符，并设置该占位符的字体、粗体和下划线。填入的金额继承占位符首个文本节点所在 run 的格式。已有表格空白单元格也可以直接用坐标定位，示例见配置说明。程序不按视觉上的下划线长度猜测位置。
6. 更新 `fields[].target` 和 `checks`，使必填项目完整、资产负债等关系覆盖真实报表；检查模板中的金额单位与 `format.unit` 一致。Word 中公司、日期、文字叙述和单位是模板静态内容，需要由你更新。
7. 运行下面的定位命令，再运行 `--dry-run`；确认日志后正式运行，并打开报告检查排版。

```sh
python -m accounting_report inspect-excel --excel your_balance.xlsx --config your_mapping.json
python -m accounting_report inspect-word --template your_report.docx
```

`inspect-word` 会列出每个占位符及其段落、正文顶层表格每个单元格的索引和原文。Word 表格坐标以 **0** 开始，依据实际 XML 单元格顺序；合并单元格会改变视觉列与实际单元格索引的关系。复杂或嵌套单元格优先使用占位符。

## 可选 AI 模块

### DeepSeek V4：识别新的 Excel 布局

新增 `inspect-layout`（本地结构扫描）、`discover`（模型定位）和 `approve-mapping`（确认并预演校验）三个命令。默认使用官方 `deepseek-v4-pro`，接口为 `https://api.deepseek.com/chat/completions`，密钥读取环境变量 `DEEPSEEK_API_KEY`。没有密钥时，原有示例和核心填报仍然可以运行。

完整步骤见 [DeepSeek 接入与布局识别](docs/deepseek.md)。模型提出科目行、金额单元格和期间表头的坐标；程序从原 Excel 读取金额。确认映射且金额/Word 预演全部通过后才保存配置。保存后的配置可以离线复用。

```powershell
python -m accounting_report inspect-layout --excel examples/balance_sheet.xlsx --sheet "资产负债表" --output results/layout.json
python -m accounting_report discover --excel examples/balance_sheet.xlsx --config examples/mapping.json --sheet "资产负债表" --output results/proposal.json
```

先按接入文档创建 `results` 目录并设置密钥。布局识别需要种子配置来定义必填科目、预期公司/期间/单位、Word 目标和勾稽关系；它不会独立推断或改变这些业务规则。换期间时先更新种子配置的 metadata 和 Word 静态日期。

### 在已配置区域内匹配科目名称

无需 API key 的本地文字相似度建议：

```sh
python -m accounting_report suggest --excel examples/balance_sheet.xlsx --config examples/mapping.json --output suggestions.json
```

这只是字符相似度排序，不是语义模型。分数不表示会计含义等价或已通过验证。

需要 AI 语义建议时，选择支持 Chat Completions 以及 JSON 对象响应的服务，显式传入完整 HTTPS 接口地址和模型名。模型不固定，避免绑定某个账户的模型权限。API key 只从环境变量读取：

```powershell
$env:REPORT_AI_API_KEY = "你的密钥"
python -m accounting_report suggest --excel examples/balance_sheet.xlsx --config examples/mapping.json --ai --endpoint "https://api.openai.com/v1/chat/completions" --model "你的可用模型名" --output ai_suggestions.json
```

DeepSeek 可以省略接口和模型参数：

```powershell
python -m accounting_report suggest --excel examples/balance_sheet.xlsx --config examples/mapping.json --ai --provider deepseek --output ai_suggestions.json
```

这条命令读取 `DEEPSEEK_API_KEY`；仍然只生成名称匹配建议。

`--ai` 会把**字段 ID、候选科目名称、分区 ID 和科目所在坐标**发送给你指定的服务；不发送金额、完整文件或 `metadata`，但科目名称本身也可能有业务敏感性。默认核心流程不会发网络请求。AI 建议必须引用已存在的字段和候选坐标，跨 block、未知或重复字段会被拒绝；它不会修改配置、生成金额或直接填报。

人工确认建议后，将真实科目名称加入对应 `source.labels`，再走本地严格验证流程。含义不同但相似的名称不要加入同一个别名列表。未匹配的必填项目仍会阻断输出。AI 测试使用模拟 HTTP 返回；本次未使用真实 key 或实测模型服务。

## 测试与重新生成示例

```sh
python -m pip install -r requirements-dev.txt
python -m unittest discover -s tests -v
```

测试覆盖端到端结果、两期金额、余额平衡、缺失和零值、重复科目、坐标防漂移、公式无缓存、Excel 错误、数字格式、单位换算、舍入、上下界、Word 跨 run 格式、页眉页脚、嵌套表格、空单元格、未知占位符、受保护文档、修订、日志、退出码、覆盖防护和 AI 返回校验。

重新生成示例 Word 和配置：

```sh
python examples/create_templates.py
```

该命令会覆盖两个示例文件，仅用于开发和恢复原样。它不会生成新的 Excel；示例 Excel 已随项目提供。实际使用时请复制模板和配置后再改，保留原始示例便于对照。

## 支持范围和需要注意的行为

- 支持普通 OOXML `.xlsx/.xlsm` 和 `.docx`；不支持扫描图片、旧 `.xls/.doc`、加密文件、OCR。签名 DOCX 修改后会使签名失效，应使用未签名模板。
- `openpyxl` 不计算 Excel 公式。读取的是保存的缓存结果；缓存缺失则失败，缓存存在也可能已过期，必须先重算保存。项目不会保证缓存新鲜性，也不会联网刷新外部链接。源金额单元格的直接外部工作簿引用被拒绝；间接外部依赖或 Excel 中人为编写的公式需在源文件中核实。
- 核心填报依据配置。可选 AI 可提出新布局，但不自行决定报表口径、币种、单位或期间。严格验证覆盖配置里的字段和规则，不表示审计了整本工作簿。完整性由必填字段清单和勾稽配置决定。
- 默认必填字段不允许空值。真实零值可以填写；只有明确设置 `blank_zero: true` 才把空值当零并记录警告。横杠 `-` 不自动当零。
- 勾稽检查可分别覆盖期初和期末，字段可同时用于核对而没有 `target`。校验使用原始数值；输出单位换算和舍入后，明细显示和合计显示可能有末位差额，不会人为调平。
- 金额类型支持数字、千位分隔文本、括号负数、全角数字、人民币符号。无效千位分组、百分号、混合单位、布尔值、科学计数文本和非有限数值会被拒绝。输入最大有效数字 28 位；Excel 本身约 15 位有效数字的数值精度限制无法由后续程序恢复。
- 只支持正文、页眉页脚中的占位符及简单正文表格单元格定位。占位符应为普通可见文本；不要放在域、公式、页码、批注、脚注或尾注中。未支持部件中的占位符会被拒绝。含修订或启用编辑保护的模板会被阻断，请先接受或拒绝修订、解除保护。
- 保留格式指不重建原文段落和表格、保留已有格式节点及未修改部件；金额文字长度变化仍可能引起 Word 重新分页。占位符不能跨段落、换行或制表符，不能把不同格式的占位符片段当作不同数字部分的样式。过长金额可能超出原空行或单元格，真实模板需打开 Word 做版面验收。
- 单元格定位仅支持正文顶层表格、没有嵌套表格且只有一个段落的单元格；用 `expected_text` 精确校验原文，不能默认覆盖已有金额。多次运行需要新的输出目录。输入文件在运行中发生变化时会停止发布。
- 核对日志含真实业务金额和文件路径，应按你的业务文件方式保管。核心项目不上传日志、不存储 API key。

## 参考文档

- [openpyxl 读取工作簿与 data_only](https://openpyxl.readthedocs.io/en/stable/tutorial.html#loading-from-a-file)：说明公式值来自上次保存的结果。
- [Python Decimal](https://docs.python.org/3/library/decimal.html)：金额计算和舍入依据。
- [OpenAI Chat Completions API](https://developers.openai.com/api/reference/resources/chat)：可选适配器的接口格式；其他兼容服务需核查其支持情况。
