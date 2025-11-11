import pandas as pd
from sentence_transformers import SentenceTransformer, util
from thefuzz import fuzz
import re
import time
import sys
import os

EXCEL_FILE_NAME = 'Outlet_Tagging.xlsx'
INPUT_FOLDER = 'Tagging Outlets'

DICTIONARY_TAB = 'Dictionary'
BRANDS_TAB = 'Brands'

DICT_SKU_COL = 'Product Name'
DICT_VISIBLE_KW_COL = 'Generic Keyword Visible'
DICT_NONVISIBLE_KW_COL = 'Generic Keyword Not Visible'
DICT_CATEGORY_COL = 'Category'
DICT_BASTYPE_COL = 'Basic Type'
TAG_SKU_COL = 'Name'
BRANDS_COL = 'Brand Name'

SEMANTIC_SCORE_WEIGHT = 0.55
KEYWORD_SCORE_WEIGHT = 0.35
BRAND_SCORE_WEIGHT = 0.10

NO_MATCH_THRESHOLD = 75

TIE_BREAKER_MARGIN = 2.5

GENERIC_IDENTIFIERS = {'bulk', 'loose', 'local'}

def preprocess_text(text):
    if not isinstance(text, str): return ""
    text = text.lower()
    text = text.replace('-', ' ')
    text = re.sub(r'[\d\.]+(kg|g|ml|l|pcs)\b', '', text)
    text = text.replace('(', ' ').replace(')', ' ')
    text = re.sub(r'sku\s*\d+', '', text)
    text = re.sub(r'[^a-z0-9\s]', '', text)
    stop_words = {'and', 'with', 'in', 'for', 'the', 'a', 'an', 'premium', 'flavoured', 'flavored'}
    query = ' '.join(word for word in text.split() if word not in stop_words)
    return re.sub(r'\s+', ' ', query).strip()

def load_brands(xls):
    try:
        df_brands = pd.read_excel(xls, sheet_name=BRANDS_TAB)
        brand_list = df_brands[BRANDS_COL].dropna().astype(str).str.lower().tolist()
        print(f"Loaded {len(brand_list)} brands successfully.")
        return set(brand_list)
    except Exception:
        print(f"Warning: Could not load '{BRANDS_TAB}' sheet. Brand logic will be limited.")
        return set()

def extract_brands(processed_text, brand_set):
    found_brands = set()
    words_in_sku = set(processed_text.split())
    for brand in brand_set:
        if set(brand.split()).issubset(words_in_sku):
            found_brands.add(brand)
    return found_brands

def calculate_brand_score(sku_brands, dict_brands, is_generic_sku):
    if is_generic_sku:
        return 100 if not dict_brands else 0
    if sku_brands and dict_brands:
        return 100 if sku_brands.intersection(dict_brands) else 0
    elif not sku_brands and not dict_brands:
        return 100
    else:
        return 50

def calculate_precise_keyword_score(sku_words, dict_words):
    if not sku_words or not dict_words: return 0
    
    mismatched_words = sku_words.difference(dict_words)
    if mismatched_words:
        penalty_factor = 1.0 - (len(mismatched_words) / len(sku_words)) * 0.75
        if penalty_factor < 0: penalty_factor = 0
    else:
        penalty_factor = 1.0

    intersection = sku_words.intersection(dict_words)
    union = sku_words.union(dict_words)
    
    base_score = (len(intersection) / len(union)) * 100
    
    return base_score * penalty_factor


if __name__ == "__main__":
    main_start_time = time.time()
    try:
        if not os.path.exists(EXCEL_FILE_NAME): raise FileNotFoundError(f"'{EXCEL_FILE_NAME}' not found.")
        if not os.path.exists(INPUT_FOLDER): os.makedirs(INPUT_FOLDER)
        
        print("--- V6.4 Definitive Intelligence: Batch Processing Mode ---")
        print("\nStep 1: Loading central resources (Excel, Brands)...")
        with pd.ExcelFile(EXCEL_FILE_NAME) as xls:
            df_dict = pd.read_excel(xls, sheet_name=DICTIONARY_TAB)
            brand_set = load_brands(xls)
        
        print("\nStep 2: Initializing ML model (this may take a moment)...")
        model = SentenceTransformer('all-MiniLM-L6-v2')
        
        print("\nStep 3: Pre-computing dictionary embeddings...")
        df_dict['processed'] = df_dict[DICT_SKU_COL].astype(str).apply(preprocess_text)
        df_dict['brands'] = df_dict['processed'].apply(lambda txt: extract_brands(txt, brand_set))
        dict_embeddings = model.encode(df_dict['processed'].tolist(), convert_to_tensor=True, show_progress_bar=True)

        print(f"\nStep 4: Discovering files in '{INPUT_FOLDER}' folder...")
        files_to_process = [f for f in os.listdir(INPUT_FOLDER) if f.endswith('.csv')]
        if not files_to_process:
            print("\nWARNING: No CSV files found in the 'Tagging Outlets' folder. Nothing to process.")
        else:
            print(f"Found {len(files_to_process)} CSV files to process.")

        for filename in files_to_process:
            file_start_time = time.time()
            filepath = os.path.join(INPUT_FOLDER, filename)
            print(f"\n--- Processing: {filename} ---")
            
            try:
                df_tag = pd.read_csv(filepath, usecols=[0], header=None, names=[TAG_SKU_COL])
                df_tag.dropna(subset=[TAG_SKU_COL], inplace=True)

                print(f"  ...Loaded {len(df_tag)} SKUs. Preprocessing...")
                df_tag['processed'] = df_tag[TAG_SKU_COL].astype(str).apply(preprocess_text)
                df_tag['brands'] = df_tag['processed'].apply(lambda txt: extract_brands(txt, brand_set))
                
                print("  ...Starting matching process...")
                results = []
                for index, row in df_tag.iterrows():
                    processed_sku = row['processed']
                    if not processed_sku: continue

                    sku_words = set(processed_sku.split())
                    is_generic_sku = bool(sku_words.intersection(GENERIC_IDENTIFIERS))
                    
                    sku_embedding = model.encode(processed_sku, convert_to_tensor=True)
                    hits = util.semantic_search(sku_embedding, dict_embeddings, top_k=10)[0]
                    
                    scored_candidates = []
                    for hit in hits:
                        dict_index = hit['corpus_id']
                        cand_row = df_dict.iloc[dict_index]
                        cand_words = set(cand_row['processed'].split())

                        semantic_score = hit['score'] * 100
                        brand_score = calculate_brand_score(row['brands'], cand_row['brands'], is_generic_sku)
                        precise_keyword_score = calculate_precise_keyword_score(sku_words, cand_words)
                        
                        hybrid_score = (semantic_score * SEMANTIC_SCORE_WEIGHT) + \
                                       (brand_score * BRAND_SCORE_WEIGHT) + \
                                       (precise_keyword_score * KEYWORD_SCORE_WEIGHT)

                        scored_candidates.append({
                            'score': hybrid_score, 'index': dict_index,
                            'intersection_count': len(sku_words.intersection(cand_words))})

                    scored_candidates.sort(key=lambda x: x['score'], reverse=True)
                    best_candidate = scored_candidates[0]
                    second_candidate = scored_candidates[1] if len(scored_candidates) > 1 else None

                    if second_candidate and abs(best_candidate['score'] - second_candidate['score']) < TIE_BREAKER_MARGIN:
                        if second_candidate['intersection_count'] > best_candidate['intersection_count']:
                            best_candidate = second_candidate
                            
                    if best_candidate['score'] < NO_MATCH_THRESHOLD:
                        debug_reason = f"REJECTED: Best score ({best_candidate['score']:.0f}%) below threshold."
                        results.append({'SKU_to_Tag': row[TAG_SKU_COL], 'Final_Score': f"{best_candidate['score']:.0f}%", 'Match_Level': 'Low (Rejected)', 'Debug_Reasoning': debug_reason})
                        continue
                    
                    best_match_row = df_dict.iloc[best_candidate['index']]
                    results.append({
                        'SKU_to_Tag': row[TAG_SKU_COL], 'Final_Score': f"{best_candidate['score']:.0f}%",
                        'Match_Level': 'High Confidence', 'Matched_Dictionary_SKU': best_match_row.get(DICT_SKU_COL, ''),
                        'Debug_Reasoning': f"Final Score: {best_candidate['score']:.0f}%"})

                print("  ...Matching complete. Saving results...")
                df_results = pd.DataFrame(results).fillna('')
                df_results = df_results[['SKU_to_Tag', 'Matched_Dictionary_SKU', 'Final_Score', 'Match_Level', 'Debug_Reasoning']]
                
                output_sheet_name = f"{os.path.splitext(filename)[0]} Result"

                with pd.ExcelWriter(EXCEL_FILE_NAME, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
                    df_results.to_excel(writer, sheet_name=output_sheet_name, index=False)
                
                file_end_time = time.time()
                print(f"  ...SUCCESS! Processed {filename} in {file_end_time - file_start_time:.2f}s. Results saved to sheet: '{output_sheet_name}'")

            except Exception as e:
                print(f"  ...ERROR processing {filename}: {e}. Skipping this file.")
                continue

        # --- PHASE 5: COMPLETION ---
        main_end_time = time.time()
        print("\n" + "="*50)
        print("          BATCH PROCESSING COMPLETE!")
        print(f"Total processing time: {main_end_time - main_start_time:.2f} seconds.")
        print(f"All results have been saved to new sheets in '{EXCEL_FILE_NAME}'.")
        print("="*50)

    except FileNotFoundError as e: print(f"\nFATAL ERROR: {e}")
    except KeyError as e: print(f"\nFATAL ERROR: A required column is missing from your Excel file: {e}")
    except Exception as e: print(f"\nAn unexpected error occurred during setup: {e}")