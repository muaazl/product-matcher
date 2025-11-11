**3. Update the Script File**

*   Open the `Code.gs` file in your editor.
*   **Delete all the old code** from version 1 that is currently in the file.
*   **Copy and paste the entire new script** (the V2 script you just provided) into the `Code.gs` file.
*   Save the file.

**4. Update the `README.md` file for this branch**

Your project's documentation needs to reflect the new features and setup. Replace the entire content of your `README.md` file with the text below, which describes V2.

```markdown
# SKU Tagger V2 (Google Apps Script)

This is the second version of the SKU Tagger, built as a Google Apps Script to run directly within a Google Sheet. It includes an improved V2 matching algorithm and retains the original V1 logic as a legacy option.

## New in Version 2

-   **Two-Stage Matching:** A high-precision first pass for exact matches, followed by a high-recall second pass using fuzzy logic for remaining items.
-   **Normalization Sheet:** A new `Normalization` sheet allows you to define custom text substitutions (e.g., "chicken flavour" -> "chicken") before the matching process begins.
-   **Improved UI:** A clearer menu in the Google Sheet UI separates the V2 and Legacy V1 taggers.
-   **Richer Debugging:** For medium-confidence matches, the output can now show the top alternative to help with manual review.

## Setup

1.  **Open your Google Sheet.**
2.  Go to **Extensions > Apps Script**.
3.  Delete any default code in the `Code.gs` file.
4.  Copy and paste the entire content of this repository's `Code.gs` file into the editor.
5.  **Save the script.**

## Sheet Naming Convention

This script requires the following sheets to exist with these exact names for V2 to work correctly:

-   `Dictionary`: Your master list of products.
-   `TaggingSheet`: The sheet containing the SKUs you want to tag.
-   `Config`: A sheet to set confidence thresholds (Cell `B1` for HIGH, `B2` for MEDIUM).
-   `Normalization`: (Optional) A sheet with two columns: Column A is the text to find, and Column B is the text to replace it with.

## How to Run

1.  Refresh your Google Sheet.
2.  A new custom menu named "⚙️ SKU Tagger V2" will appear.
3.  Click the menu and choose from the options:
    -   **🚀 Run Tagger V2.1 (NEW):** Executes the new, more advanced two-stage matching algorithm.
    -   **Run Legacy Tagger V1.1:** Executes the original matching algorithm from Version 1.
4.  The script will run and place the results in the `TaggingSheet`, starting from column B.