# AI CellFill Excel

此脚本使用多个AI API（在`.config`中配置）来自动将内容填充到Excel电子表格中，根据用户提示。

## 功能

- 从Excel文件（`prompts.xlsx`）中读取用户提示。
- 允许使用标志列选择性生成。
- 从外部文件（`.config`和`systemprompt.txt`）读取API配置和系统提示。
- 调用指定的AI API根据提示生成文本。
- 将生成的文本写回Excel文件。
- 使用`openpyxl`对Excel文件应用基本格式（列宽、字体样式、文本换行）。
- 在每次成功的API响应后保存Excel文件以防止数据丢失。
- 使用`.config`文件来确定使用哪些AI模型，检查`ENABLED`标志。

## 设置

1. **克隆/下载：** 获取脚本文件（`main.py`、`.config`、`systemprompt.txt`、`fonts/`目录）。
2. **安装依赖：** 你需要Python 3和以下库。在项目目录中打开终端并运行：

    ```bash
    pip install requests openpyxl
    ```

3. **配置API：**
    - 编辑`.config`文件。
    - 为你想要使用的每个AI模型添加以`[API_]`开头的部分。
    - 将每个模型的占位符`KEY`替换为你的实际API密钥。
    - 确保每个API的`ENDPOINT`、`MODEL`、`NAME`、`TYPE`和`ENABLED`设置正确。

    ```ini
    [API_X_AI]
    KEY = your_x.ai_api_key_here
    ENDPOINT = https://api.x.ai/v1
    MODEL = grok-2-latest
    NAME = X_AI
    TYPE = openai
    ENABLED = true

    [API_ALIYUN_QWEN]
    KEY = your_aliyun_qwen_api_key_here
    ENDPOINT = https://qwen.aliyun.com/v1
    MODEL = qwen-1.5
    NAME = Aliyun_Qwen
    TYPE = openai
    ENABLED = true

    [API_GOOGLE_GEMINI]
    KEY = your_google_gemini_api_key_here
    ENDPOINT = https://generativelanguage.googleapis.com
    MODEL = gemini-pro
    NAME = Google_Gemini
    TYPE = google
    ENABLED = true
    ```

4. **系统提示（可选）：**
    - 编辑`systemprompt.txt`文件。
    - 添加你希望AI对所有提示遵循的任何系统级指令。如果不需要系统提示，请将其留空。
5. **字体：**
    - 脚本尝试使用`main.py`中指定的`TARGET_FONT_NAME`字体（默认：'Source Han Sans VF'）。
    - 确保字体文件（`SourceHanSans-VF.ttf.ttc`）位于`fonts`目录中。
    - **重要：** `openpyxl`应用字体*名称*。你必须在查看Excel文件的系统上安装字体才能正确渲染。脚本本身不嵌入字体文件。

## 使用

1. **准备Excel文件（`prompts.xlsx`）：**
    - 如果文件不存在，脚本将创建它并添加必要的标题。
    - **列A（用户提示词）：** 从第2行开始，每行输入一个提示。
    - **列B（是否生成(0是1否))：** 输入`0`表示你希望脚本处理并生成文本的行。输入`1`（或留空/使用其他值）表示你希望跳过的行。
    - **列C起：** 这些列将由脚本填充每个启用模型的AI响应。对于标记为`0`的行，这些列中的现有内容将被覆盖。
2. **运行脚本：** 在项目目录中打开终端并运行：

    ```bash
    python main.py
    ```

3. **检查输出：**
    - 脚本将在终端中打印进度消息，包括API调用和保存操作。
    - 完成后，打开`prompts.xlsx`查看结果。对于标记为`0`的行，生成的文本将在对应启用AI模型的列中。应应用格式。

## 注意事项

- 脚本在每次成功的API调用和生成后保存Excel文件。这是更安全的，但对于许多提示可能会更慢。
- 使用`openpyxl`对Excel进行格式化有限制，特别是行高自动调整。你可能需要在Excel中手动调整格式以获得完美的外观。
- 确保你的API密钥保密，不公开共享。
