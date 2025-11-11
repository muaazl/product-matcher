# SKU Tagger V1 (Google Apps Script)

This is the first version of the SKU Tagger, built as a Google Apps Script to run directly within a Google Sheet.

## Features

-   Pure JavaScript implementation (Google Apps Script).
-   Runs directly inside the Google Sheets environment.
-   Uses Sorensen-Dice and Levenshtein distance for string matching.
-   Adds a custom menu item to the spreadsheet UI to trigger the tagging process.
-   Color-codes the output rows based on confidence scores (High, Medium, Low).

## Setup

1.  **Open your Google Sheet.**
2.  Go to **Extensions > Apps Script**.
3.  Delete any default code in the `Code.gs` file.
4.  Copy and paste the entire content of this repository's `Code.gs` file into the editor.
5.  **Save the script.**

## Sheet Naming Convention

This script requires the following sheets to exist with these exact names:

-   `Dictionary`: Your master list of products.
-   `SPAR (Gampaha)`: The sheet containing the SKUs you want to tag. (Note: This is hard-coded and must be changed in the script if you use a different sheet name).
-   `Config`: A sheet to set confidence thresholds. (Cell `B1` for HIGH, `B2` for MEDIUM).

## How to Run

1.  Refresh your Google Sheet.
2.  A new custom menu named "SKU Tagger V1" will appear.
3.  Click on **SKU Tagger V1 > Start Tagging Process**.
4.  The script will run and place the results in the `SPAR (Gampaha)` sheet, starting from column B.