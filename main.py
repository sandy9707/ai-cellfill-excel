import os
import json  # Needed for Google API
from utils.config import read_config
from utils.api import call_api
from utils.excel import (
    initialize_excel,
    apply_formatting,
    find_column_index,
    read_and_process_excel,
    write_excel_with_defaults,
)
from utils.system_prompt import read_system_prompt

# --- Configuration ---
CONFIG_FILE = ".config"  # Path to the configuration file storing API keys
EXCEL_FILE = "prompts.xlsx"  # Path to the Excel file for prompts and results
SYSTEM_PROMPT_FILE = (
    "systemprompt.txt"  # Path to the file containing the system prompt for the AI
)
FONT_PATH = "fonts/hei.TTF"  # Path to the desired font file (used for reference)
TARGET_FONT_NAME = "hei"  # Font name as recognized by Excel
API_TIMEOUT = 180  # Increased timeout for potentially multiple long calls

# --- Main Execution ---


def main():
    """
    Main function to orchestrate the script's execution.
    1. Reads all LLM configurations.
    2. Reads system prompt.
    3. Initializes or loads the Excel file, matching headers to configs.
    4. Iterates through rows, calling all configured APIs for rows marked '0'.
    5. Saves results to respective columns and updates the flag.
    6. Applies formatting and saves periodically and at the end.
    """
    print("Starting AI CellFill Excel script (Multi-LLM Version)...")

    # 1. Read Configs
    llm_configs = read_config(CONFIG_FILE)
    if not llm_configs:
        print("Exiting due to configuration errors.")
        return  # Stop if config is invalid or empty

    # 2. Read System Prompt
    system_prompt = read_system_prompt(SYSTEM_PROMPT_FILE)
    if system_prompt:
        print(f"Using System Prompt (first 50 chars): '{system_prompt[:50]}...'")
    else:
        print("No system prompt found or read.")

    # 3. Initialize/Load Excel & Get Column Mapping
    workbook, sheet, llm_col_map = initialize_excel(llm_configs, EXCEL_FILE)
    if not workbook or not sheet or not llm_col_map:
        print("Exiting due to Excel initialization errors.")
        return  # Stop if Excel handling fails

    # Update first column header to "用户指南"
    sheet.cell(row=1, column=1).value = "用户指南"

    # Populate first few rows with specified information
    if sheet.max_row < 5:  # Ensure there are enough rows
        for i in range(sheet.max_row + 1, 6):
            sheet.cell(row=i, column=1)  # Add empty cells if needed

    sheet.cell(row=2, column=1).value = "系统提示词"
    sheet.cell(row=3, column=1).value = system_prompt  # System prompt from file
    sheet.cell(row=4, column=1).value = "本列自动生成, 仅供展示"
    sheet.cell(row=5, column=1).value = "已启用模型为:"
    enabled_models = "; ".join(
        [config["NAME"] for config in llm_configs if config["ENABLED"]]
    )
    sheet.cell(row=6, column=1).value = f"启用模型: {enabled_models}"

    # Find the flag column index dynamically
    flag_header = "是否生成 (0 是 1 否)"
    flag_col_idx = find_column_index(sheet, flag_header)
    if not flag_col_idx:
        print(f"Critical Error: Cannot find flag column '{flag_header}'. Exiting.")
        return

    # Ensure '是否生成 (0 是 1 否)' column has defaults (run once before processing)
    read_and_process_excel(EXCEL_FILE)

    # 4. Process Rows in Excel File
    rows_processed_count = 0
    print("Processing Excel rows...")
    for row_idx in range(2, sheet.max_row + 1):  # Start from row 2 (skip header)
        row_needs_processing = False
        try:
            # Get cells for the current row
            should_generate_cell = sheet.cell(row=row_idx, column=flag_col_idx)
            user_prompt_cell = sheet.cell(
                row=row_idx, column=2
            )  # Prompt is now in column 2

            # Check if generation is requested (Flag Column value is 0)
            should_generate_val = should_generate_cell.value
            # Handle potential non-numeric values gracefully
            try:
                if int(should_generate_val) == 0:
                    row_needs_processing = True
                elif int(should_generate_val) == 1:
                    # Skip if Column B is explicitly set to 1
                    # print(f"Skipping row {row_idx}: Marked as '1' (Do not generate).") # Reduce verbosity
                    pass
                else:
                    print(
                        f"Skipping row {row_idx}: Invalid value '{should_generate_val}' in '{flag_header}' column (must be 0 or 1)."
                    )
            except (ValueError, TypeError):
                print(
                    f"Skipping row {row_idx}: Non-numeric value '{should_generate_val}' in '{flag_header}' column. Setting to 0."
                )
                should_generate_cell.value = 0  # Default to 0 if invalid
                row_needs_processing = True

            if row_needs_processing:
                user_prompt = (
                    str(user_prompt_cell.value).strip()
                    if user_prompt_cell.value
                    else None
                )

                if user_prompt:
                    print(
                        f"Processing row {row_idx} for prompt: '{user_prompt[:50]}...'"
                    )
                    all_apis_called_for_row = True  # Assume success until an API fails

                    # Iterate through each configured LLM
                    for api_config in llm_configs:
                        if not api_config["ENABLED"]:
                            continue  # Skip disabled LLMs
                        llm_name = api_config["NAME"]
                        output_col_idx = llm_col_map.get(llm_name)

                        if not output_col_idx:
                            print(
                                f"  Error: Could not find column index for LLM '{llm_name}'. Skipping API call."
                            )
                            all_apis_called_for_row = (
                                False  # Mark row as incomplete if column missing
                            )
                            continue

                        output_cell = sheet.cell(row=row_idx, column=output_col_idx)

                        # Optional: Check if this specific cell already has content?
                        # if output_cell.value:
                        #     print(f"  Skipping API call for '{llm_name}' - cell already has content.")
                        #     continue

                        api_result = call_api(api_config, system_prompt, user_prompt)

                        # Write result (or error message) to the LLM's column without metadata
                        output_cell.value = (
                            api_result if api_result else "No result returned."
                        )
                        if api_result and api_result.startswith("Error:"):
                            print(
                                f"  API call failed for '{llm_name}' (Row {row_idx}). Error written to cell."
                            )
                            # Decide if one failure should stop flag update? Current logic updates flag anyway.
                            # all_apis_called_for_row = False # Uncomment this if one failure should prevent flag update
                        else:
                            print(
                                f"  Result for '{llm_name}' written to row {row_idx}, column {output_col_idx}."
                            )

                    # After trying all APIs for this row, update the flag if all were attempted
                    # (Modify condition 'all_apis_called_for_row' if needed based on failure handling)
                    if all_apis_called_for_row:
                        should_generate_cell.value = 1
                        print(f"Flag updated to 1 for row {row_idx}.")
                        print(
                            f"  Metadata should have been added for processed APIs in row {row_idx}."
                        )
                        rows_processed_count += 1
                    else:
                        print(
                            f"Flag *not* updated for row {row_idx} due to previous errors or skipped calls."
                        )

                    # Save after each *row* is fully processed to prevent data loss
                    try:
                        # Apply formatting before saving (optional, can be slow)
                        # apply_formatting(sheet, len(llm_configs))
                        workbook.save(EXCEL_FILE)
                        print(f"'{EXCEL_FILE}' saved after processing row {row_idx}.")
                    except Exception as e:
                        print(
                            f"Error saving '{EXCEL_FILE}' after processing row {row_idx}: {e}"
                        )
                        print("Attempting to continue...")

                else:
                    if not user_prompt:
                        print(f"Skipping row {row_idx}: Empty user prompt.")
                    else:
                        print(
                            f"Skipping row {row_idx}: Row is within first 6 display rows."
                        )
                    if row_idx > 6:
                        should_generate_cell.value = (
                            1  # Mark as done only if beyond first 6 rows
                        )
                        print(
                            f"Flag updated to 1 for row {row_idx} (empty prompt or display row)."
                        )
                    continue

        except (
            Exception
        ) as e:  # Catch any other unexpected errors during row processing
            print(f"Critical Error processing row {row_idx}: {e}")
            # Consider logging traceback here
            # import traceback
            # traceback.print_exc()
            # Try to mark the row with an error message in the first LLM column if possible
            try:
                first_llm_col = min(llm_col_map.values())
                sheet.cell(row=row_idx, column=first_llm_col).value = (
                    f"Error processing row: {e}"
                )
                # Also mark as 'done' to avoid retrying a broken row
                if flag_col_idx:
                    sheet.cell(row=row_idx, column=flag_col_idx).value = 1
            except:
                pass  # Ignore if we can't even write the error

    # 5. Final Save and Formatting Application
    try:
        print("\nApplying final formatting...")
        apply_formatting(
            sheet, len(llm_configs)
        )  # Ensure formatting is applied one last time
        workbook.save(EXCEL_FILE)
        print(f"Final save of '{EXCEL_FILE}' complete.")
    except Exception as e:
        print(f"Error during final save/formatting: {e}")

    print(
        f"\nScript finished. Processed {rows_processed_count} rows for all configured LLMs."
    )


if __name__ == "__main__":
    main()
