# AI CellFill Excel

一个使用AI API在Excel中自动生成内容的Python脚本。该工具将用户提示与系统提示结合，通过配置的语言模型API生成内容。

## 功能

- **多LLM支持**：配置多个语言学习模型（LLM）来生成内容。
- **Excel集成**：读取和写入Excel文件，用于管理提示和结果。
- **自定义系统提示**：使用自定义系统提示进行内容生成。
- **条件处理**：仅处理用户提示非空且超出初始显示行的行。

## 文件结构

- **main.py**：运行应用程序的主脚本。
- **utils/api.py**：处理对配置的LLM的API调用。
- **utils/config.py**：管理配置设置和API密钥。
- **utils/excel.py**：管理Excel文件操作。
- **utils/system_prompt.py**：从文本文件读取系统提示。
- **prompts.xlsx**：用于输入提示和输出结果的Excel文件。
- **systemprompt.txt**：包含AI系统提示的文本文件。

## Excel文件结构

Excel文件 `prompts.xlsx` 的结构如下：
- **第1列 - 用户指南**：显示列，包含系统提示和配置详细信息（不用于生成）。
- **第2列 - 用户提示词**：包含与系统提示结合用于生成的用戶提示。
- **第3列 - 是否生成 (0 是 1 否)**：指示是否应处理该行的标志（0表示是，1表示否）。
- **第4列 - X_AI**：X_AI模型生成的输出内容列。

## 使用方法

1. **配置**：编辑 `.config` 文件以设置API密钥并启用/禁用LLM。
2. **系统提示**：更新 `systemprompt.txt` 以设置所需的系统提示。
3. **Excel设置**：在 `prompts.xlsx` 的"用户提示词"列中添加用户提示，并将"是否生成 (0 是 1 否)"标志设置为0以处理这些行。
4. **运行脚本**：执行 `python main.py` 以处理提示并生成内容。

## 要求

- Python 3.x
- 库：`openpyxl`，`requests`（通过 `pip install openpyxl requests` 安装）

## 许可证

该项目采用MIT许可证 - 有关详细信息，请参见LICENSE文件。
