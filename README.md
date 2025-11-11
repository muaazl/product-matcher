# Definitive Intelligence: SKU Matching Engine

This project uses a hybrid AI model to intelligently match product names (SKUs) from various sources against a master dictionary. It is ideal for data cleaning, product tagging, and catalog harmonization in e-commerce or retail analytics.

## Features

-   **Hybrid Scoring Model:** Combines multiple techniques for robust matching:
    -   **Semantic Search:** Uses sentence-transformers (`all-MiniLM-L6-v2`) to understand the contextual meaning of product names.
    -   **Precise Keyword Analysis:** Calculates a score based on the intersection and union of keywords.
    -   **Brand Intelligence:** Identifies brands to weigh matches and correctly handles generic vs. branded products.
-   **Batch Processing:** Automatically processes multiple `.csv` files from an input folder, saving results neatly into separate Excel sheets.
-   **Customizable:** Key parameters like scoring weights and match thresholds can be easily tweaked in the script.
-   **Resilient:** Handles missing data and ignores irrelevant columns in input files.

## Project Structure

Your project folder should be set up as follows for the script to run correctly:

## Setup & Installation

1.  **Clone the repository:**
    ```sh
    git clone https://github.com/muaazl/product-matcher.git
    cd product-matcher
    ```

2.  **Create and activate a virtual environment (highly recommended):**
    ```sh
    # For Windows
    python -m venv venv
    .\venv\Scripts\activate

    # For macOS/Linux
    python3 -m venv venv
    source venv/bin/activate
    ```

3.  **Install the required dependencies:**
    ```sh
    pip install -r requirements.txt
    ```

## How to Run

1.  **Prepare Your Data:**
    *   Update `Outlet_Tagging.xlsx` with your master `Dictionary` and `Brands` sheets.
    *   Place one or more `.csv` files containing the product SKUs you want to tag into the `Tagging Outlets` folder. **The script will only read the first column (Column A) of each CSV.**

2.  **Execute the Script:**
    *   Run the main python script from your terminal:
    ```sh
    python product_matcher.py
    ```

3.  **Check the Results:**
    *   The script will process each CSV file from the input folder.
    *   For each file (e.g., `outlet_products.csv`), a new sheet named `Results_outlet_products` will be created in your `Outlet_Tagging.xlsx` file containing the matching results. The original sheets are preserved.

## Customization

The script's matching behavior can be fine-tuned by adjusting the constant variables at the top of the `product_matcher.py` file:

-   `SEMANTIC_SCORE_WEIGHT`, `KEYWORD_SCORE_WEIGHT`, `BRAND_SCORE_WEIGHT`: Adjust how much each component contributes to the final score.
-   `NO_MATCH_THRESHOLD`: The minimum score required to consider a match valid.
-   `TIE_BREAKER_MARGIN`: The score difference within which the keyword intersection count is used to break a tie.
