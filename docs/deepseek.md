# DeepSeek 接入与 Excel 布局识别

默认接入官方 DeepSeek API。截至 2026-10-09，官方接口接受模型名 `deepseek-v4-pro`；另一可选模型名为 `deepseek-flash`。界面里的“DeepSeek V4”不是直接传入接口的模型 ID。

- [官方入门与模型名称](https://api-docs.deepseek.com/en/)
- [Chat Completions 接口](https://api-docs.deepseek.com/api/create-chat-completion/)
- [JSON 输出](https://api-docs.deepseek.com/guides/json_mode/)

无需安装额外 SDK。使用项目现有依赖和 Python 标准库发起 HTTPS 请求。官方预设使用 JSON 输出、非流式和非思考模式，输出最多 8192 tokens；默认超时 120 秒，不自动重试或切换模型。错误、截断和空响应会停止处理，HTTP 响应正文不写入日志。

## 1. 在本机设置 key

在项目根目录 PowerShell 中执行。输入会隐藏，密钥只进入当前 PowerShell 及子进程的环境，不写入项目文件。不要把真实 key 发到聊天或 GitHub。

```powershell
$secret = Read-Host "输入 DeepSeek API key" -AsSecureString
$env:DEEPSEEK_API_KEY = [System.Net.NetworkCredential]::new("", $secret).Password
New-Item -ItemType Directory -Force results | Out-Null
```

项目环境中的 `python` 可以替换为 `.\.venv\Scripts\python.exe`。先依照 README 安装依赖。

## 2. 先查看本地扫描结果

```powershell
python -m accounting_report inspect-layout --excel examples/balance_sheet.xlsx --sheet "资产负债表" --range A1:AZ200 --output results/layout.json
```

此步骤不联网。查看 `layout.json`，确认范围包含公司、单位、日期表头以及全部需要的科目。默认只扫描 `A1:AZ200`；最多每张表 30000 格、8 张工作表。超过限制会报错，不会静默截断；可用 `--sheet` 限定工作表，用 `--range` 调整区域。支持双栏和合并表头的结构展示，科目文字必须在真实保存值的左上角。

扫描结果隐藏数值单元格、能解析为金额的纯文本金额和公式内容，只保留它们的类型。日期文字、科目名称、工作表名称、普通文本、公司名称、合并范围及坐标会保留。普通文本仍可能包含业务敏感信息或嵌入的数字；在发起下一步前查看本地结果。

## 3. 让模型提出新布局映射

```powershell
python -m accounting_report discover --excel examples/balance_sheet.xlsx --config examples/mapping.json --sheet "资产负债表" --range A1:AZ200 --output results/proposal.json
```

该命令显式调用 DeepSeek，传送上述扫描信息及种子配置中的字段名称、预期来源标识和期间表头，不发送完整文件、Word 内容或金额数值。根据所用模型的服务规则计费。若需要其他模型或中转服务：

```powershell
python -m accounting_report discover --excel examples/balance_sheet.xlsx --config examples/mapping.json --endpoint "https://你的服务/chat/completions" --model "服务提供的模型ID" --key-env DEEPSEEK_API_KEY --output results/proposal.json
```

自定义服务需要支持 Chat Completions 和 JSON 对象响应。程序拒绝重定向，不把密钥转发到重定向地址。官方 DeepSeek 主机专用的 `thinking` 参数不会发送给其他主机。

种子配置的职责是定义会计规则：必填字段、科目含义、期间表头、公司、单位、Word 填入目标和勾稽检查。它的旧行列可以与新表不同，但科目来源必须使用 `block/period`，每个期间列必须在 metadata 中配置实际表头。程序取原科目范围上方该期间列最近的 metadata 作为预期表头。至少有公司/单位等其他来源标识时也应保留这些检查。数值形式的公司编码或单位编码目前不能用作 AI 布局来源标识，需要文字标识。

模型必须给出真实扫描坐标、科目判断依据和原来源标识的新位置。程序拒绝未知/重复/缺失字段、同一金额格的重复读取、不同期间指向不同科目行、跨行金额、期间错位以及公司/单位/日期不符。返回不确定项时停止生成配置。程序只能验证结构，不能证明模型判断的科目语义正确；例如明细和合计的口径仍需要人工核对。勾稽通过也不能代替这一步。

## 4. 核对并保存配置

打开 `results/proposal.json`，核对 `answer.matches` 中的 `field`、`label_cell`、`value_cell`、`header_cell`、`reason`，同时检查 Word 目标与种子配置的会计规则。不要手动改 `candidate_config`；发现错误应修正种子配置或扫描范围后重新生成。以下 `--accept` 表示你已核对并接受科目语义和位置。

```powershell
python -m accounting_report approve-mapping --excel examples/balance_sheet.xlsx --config examples/mapping.json --proposal results/proposal.json --template examples/report_template.docx --output results/approved-mapping.json --output-dir results/mapping-preflight --accept
```

这一步不联网，不生成报告。它重新验证文件哈希和建议，读取真实金额，执行全部勾稽、金额范围及 Word 目标预演。通过后保留 `approved-mapping.json`，失败时移除本次新建的候选配置并保留失败核对日志。配置与输出目录必须是新路径，不能覆盖旧结果。修改源表或种子配置后，旧建议会被拒绝。

## 5. 生成报告并复用

```powershell
python -m accounting_report run --excel examples/balance_sheet.xlsx --template examples/report_template.docx --config results/approved-mapping.json --output-dir results/deepseek-report
```

此步骤完全在本地运行。确认后的配置使用固定坐标及真实科目文字保护位置；可复用到布局、科目、公司、单位和期间不变的同类表。插行、移动或改期间后需要重新配置或识别；程序不会后台调用 AI 修复。新期间也须同步 Word 模板中的静态日期。

## 测试范围

新测试用模拟 DeepSeek 响应覆盖整体行列移动后的识别→确认→预演→报告、金额隐藏、坐标和期间校验、失败日志和配置回滚。未配置用户 key，未进行真实付费请求。真实模型是否识别准确，需要用你的模板进行首次试跑。
