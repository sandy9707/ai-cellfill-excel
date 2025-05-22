# AI CellFill Excel

A Python script to automate content generation in Excel using AI APIs. This tool processes user prompts combined with a system prompt to generate content via configured Language Model APIs.

## Features

- **Multi-LLM Support**: Configure multiple Language Learning Models (LLMs) to generate content.
- **Excel Integration**: Reads from and writes to an Excel file for managing prompts and results.
- **Custom System Prompt**: Utilizes a custom system prompt for content generation.
- **Conditional Processing**: Only processes rows with non-empty user prompts and beyond the initial display rows.

## File Structure

- **main.py**: Main script to run the application.
- **utils/api.py**: Handles API calls to the configured LLMs.
- **utils/config.py**: Manages configuration settings and API keys.
- **utils/excel.py**: Manages Excel file operations.
- **utils/system_prompt.py**: Reads the system prompt from a text file.
- **prompts.xlsx**: Excel file for input prompts and output results.
- **systemprompt.txt**: Text file containing the system prompt for AI.

## Excel File Structure

The Excel file `prompts.xlsx` is structured as follows:
- **Column 1 - 用户指南 (User Guide)**: Display column with system prompt and configuration details (not used for generation).
- **Column 2 - 用户提示词 (User Prompt)**: Contains user prompts to be combined with the system prompt for generation.
- **Column 3 - 是否生成 (0 是 1 否) (Generate Flag)**: Flag to indicate if the row should be processed (0 for yes, 1 for no).
- **Column 4 - X_AI**: Output column for generated content from the X_AI model.

## Usage

1. **Configuration**: Edit `.config` file to set up API keys and enable/disable LLMs.
2. **System Prompt**: Update `systemprompt.txt` with the desired system prompt.
3. **Excel Setup**: Add user prompts in the "用户提示词" column of `prompts.xlsx` and set the "是否生成 (0 是 1 否)" flag to 0 for rows you want to process.
4. **Run Script**: Execute `python main.py` to process the prompts and generate content.

## Requirements

- Python 3.x
- Libraries: `openpyxl`, `requests` (install via `pip install openpyxl requests`)

## License

This project is licensed under the MIT License - see the LICENSE file for details.
