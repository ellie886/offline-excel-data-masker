# 离线 Excel 数据脱敏工具

[![Tests](https://github.com/ellie886/offline-excel-data-masker/actions/workflows/tests.yml/badge.svg)](https://github.com/ellie886/offline-excel-data-masker/actions/workflows/tests.yml)
[![Python 3.12](https://img.shields.io/badge/Python-3.12-3776AB.svg)](https://www.python.org/)
[![License: MIT](https://img.shields.io/badge/License-MIT-yellow.svg)](LICENSE)

面向投行、审计、财务和尽调场景的本地 Excel 数据脱敏工具。文件只在用户电脑上处理，不调用云端 API，不覆盖原文件。

![应用导入页面](docs/images/app-preview.png)

## 项目背景

财务和尽调工作簿通常包含客户、供应商、人员身份及交易信息，无法直接上传第三方平台。本项目提供网页式本地操作界面，在保留 Excel 公式和结构的前提下完成敏感信息识别、人工确认、脱敏及校验。

项目采用 AI 辅助开发方式，但应用运行时不依赖大模型或在线 AI 服务。

## 核心能力

- **离线运行：** Streamlit 仅监听 `127.0.0.1`，关闭使用统计，不调用云端 API。
- **识别与确认：** 支持手机号、身份证号、邮箱、指定列和自定义敏感词，所有候选项必须先预览再处理。
- **一致性脱敏：** 同一客户、供应商或人员在不同工作表中保持相同替换编号。
- **结构保护：** 基于 Excel 底层文件格式定点修改目标单元格，不使用普通工作簿整体重写流程。
- **自动校验：** 比较公式、数值、样式、工作表结构、隐藏状态和非目标 OOXML 组件。
- **审计报告：** 输出处理明细、风险提示和校验结果，不记录完整原始敏感值。

## 处理流程

```text
导入文件 → 设置规则 → 安全预览 → 人工确认 → 定点修改 → 自动校验 → 导出文件与报告
```

若关键校验失败，界面不会提供脱敏工作簿下载，只提供失败校验报告。

## 支持范围

当前 MVP 支持：

- 单个未加密、无宏、无数字签名的 `.xlsx` 文件；
- 手机号、18 位身份证号、邮箱和文本型银行账号；
- 指定列中的客户、供应商、人员、企业、地址和合同编号；
- 掩码、一致性编号和自定义替换；
- 隐藏工作表、隐藏行列及普通批注的风险提示；
- 公式、数值常量、样式、合并区域、冻结窗格、筛选和打印设置校验。

暂不支持：

- `.xls`、`.xlsm`、`.xlsb`、CSV；
- 加密、密码保护或带数字签名的文件；
- 金额、数量、单价、税额和比例的修改；
- 文本框、形状、图表、线程批注、嵌入对象及外部连接内容的脱敏；
- 通用本地 NER、批量处理和桌面安装包。

## 快速开始

推荐使用 Python 3.12 和 [uv](https://docs.astral.sh/uv/)：

```bash
git clone https://github.com/ellie886/offline-excel-data-masker.git
cd offline-excel-data-masker
uv venv --python 3.12
uv sync --extra dev
uv run streamlit run app.py
```

浏览器访问 `http://127.0.0.1:8501`。

依赖安装完成后，扫描、脱敏、校验和导出流程可以在断网状态下运行。不要将应用改为监听 `0.0.0.0`，也不要部署到云端后处理真实敏感文件。

## 测试

```bash
uv run pytest
uv run pytest --cov=offline_masker --cov-report=term
```

当前测试覆盖规则识别、一致性映射、共享字符串、隐藏工作表、公式、数值、样式、合并区域、审计报告及完整 Streamlit 操作流程。测试工作簿均在内存中使用虚构数据生成。

## 设计说明

`.xlsx` 本质上是一个包含 XML 文件及资源的 ZIP 包。项目只修改用户确认的目标单元格，并原样复制其他内部组件。处理完成后重新读取输出文件，并执行两层校验：

1. 对目标工作表进行结构级比较，确认除目标值外的 XML 结构一致；
2. 对其他 OOXML 组件进行字节级比较，确认未授权内容没有变化。

这可以降低整体重写工作簿造成公式、格式或高级 Excel 对象损坏的风险，但不能替代 Microsoft Excel 的完整兼容性测试。

## 安全边界

本项目不会主动上传文件，但“本地运行”不等于零风险。输出文件仍可能在当前未扫描的图表、形状、批注、嵌入对象或外部连接中保留敏感内容。浏览器扩展、本机恶意软件、云同步目录、系统交换空间以及 Excel 打开外部连接也属于风险来源。

处理真实敏感数据前，请阅读 [SECURITY.md](SECURITY.md)，使用隔离环境进行验证，并检查输出文件的全部隐藏内容。当前版本尚未经过第三方安全审计，不应直接视为金融机构生产级安全产品。

## 项目结构

```text
app.py                         Streamlit 界面
src/offline_masker/detectors  敏感信息检测
src/offline_masker/maskers    脱敏和一致性映射
src/offline_masker/excel      工作簿检查与定点修改
src/offline_masker/validation 处理后校验
src/offline_masker/reports    审计报告
tests                         单元、端到端及界面测试
```

## License

[MIT License](LICENSE)
