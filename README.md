# AI CellFill Excel

This script uses multiple AI APIs (configured in `.config`) to automatically fill content into an Excel spreadsheet based on user prompts.

## Features

- Reads user prompts from an Excel file (`prompts.xlsx`).
- Allows selective generation using a flag column.
- Reads API configuration and system prompt from external files (`.config` and `systemprompt.txt`).
- Calls the specified AI APIs to generate text based on prompts.
- Writes the generated text back into the Excel file.
- Applies basic formatting (column widths, font styles, text wrapping) to the Excel file using `openpyxl`.
- Saves the Excel file after each successful API response to prevent data loss.
- Uses the `.config` file to determine which AI models to use, checking the `ENABLED` flag.

## Setup

1. **Clone/Download:** Get the script files (`main.py`, `.config`, `systemprompt.txt`, `fonts/` directory).
2. **Install Dependencies:** You need Python 3 and the following libraries. Open your terminal in the project directory and run:

    ```bash
    pip install requests openpyxl
    ```

3. **Configure API:**
    - Edit the `.config` file.
    - Add sections for each AI model you want to use, starting with `[API_]`.
    - Replace the placeholder `KEY` with your actual API key for each model.
    - Ensure `ENDPOINT`, `MODEL`, `NAME`, `TYPE`, and `ENABLED` are correctly set for each API.

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

4. **System Prompt (Optional):**
    - Edit the `systemprompt.txt` file.
    - Add any system-level instructions you want the AI to follow for all prompts. Leave it empty if you don't need a system prompt.
5. **Font:**
    - The script attempts to use the font specified by `TARGET_FONT_NAME` in `main.py` (default: 'Source Han Sans VF').
    - Ensure the font file (`SourceHanSans-VF.ttf.ttc`) is in the `fonts` directory.
    - **Important:** `openpyxl` applies font *names*. You must have the font installed on the system where you *view* the Excel file for it to render correctly. The script itself doesn't embed the font file.

## Usage

1. **Prepare Excel File (`prompts.xlsx`):**
    - If the file doesn't exist, the script will create it with the necessary headers.
    - **Column A (用户提示词):** Enter your prompts, one per row, starting from row 2.
    - **Column B (是否生成(0是1否)):** Enter `0` for rows you want the script to process and generate text for. Enter `1` (or leave blank/use other values) for rows you want to skip.
    - **Column C onwards:** These columns will be filled by the script with the AI's responses for each enabled model. Existing content in these columns for rows marked with `0` will be overwritten.
2. **Run the Script:** Open your terminal in the project directory and run:

    ```bash
    python main.py
    ```

3. **Check Output:**
    - The script will print progress messages to the terminal, including API calls and save operations.
    - Once finished, open `prompts.xlsx` to view the results. The generated text will be in the columns corresponding to the enabled AI models for the rows you marked with `0`. Formatting should be applied.

## Notes

- The script saves the Excel file after *each* successful API call and generation. This is safer but can be slower for many prompts.
- Excel formatting with `openpyxl` has limitations, especially with row height auto-adjustment. You might need to manually adjust formatting in Excel for perfect appearance.
- Ensure your API keys are kept secure and not shared publicly.
