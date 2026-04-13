import argparse
import os
import re

import camelot
import ocrmypdf
import pandas as pd


def parse_args():
    parser = argparse.ArgumentParser(
        description="Combine scanned statement PDFs into an Excel file (Transactions only)."
    )
    parser.add_argument("--input", type=str, default=".", help="Input folder containing PDF files")
    parser.add_argument("--output", type=str, default="combined.xlsx", help="Output Excel filename")
    return parser.parse_args()


def _norm(s: str) -> str:
    # normalize spacing/case so headers match even if OCR adds extra spaces/newlines
    s = "" if s is None else str(s)
    s = s.replace("\n", " ")
    s = re.sub(r"\s+", " ", s).strip().lower()
    return s


REQUIRED_COLS = {"date", "description", "amount", "balance"}


def _looks_like_transactions_header(row_values) -> bool:
    cols = {_norm(v) for v in row_values if _norm(v)}
    return REQUIRED_COLS.issubset(cols)


def extract_transactions_tables(input_folder: str) -> pd.DataFrame:
    combined = []

    pdf_files = sorted([f for f in os.listdir(input_folder) if f.lower().endswith(".pdf")])
    cache_dir = os.path.join(input_folder, "cache")
    os.makedirs(cache_dir, exist_ok=True)

    for pdf_file in pdf_files:
        pdf_path = os.path.join(input_folder, pdf_file)
        cache_pdf_path = os.path.join(cache_dir, pdf_file)

        # OCR the PDF (skip if it was already OCR'd into cache)
        if not os.path.exists(cache_pdf_path):
            ocrmypdf.ocr(pdf_path, cache_pdf_path)

        # Extract tables
        tables = camelot.read_pdf(cache_pdf_path, pages="all", flavor="stream")

        for i, table in enumerate(tables):
            df = table.df
            if df is None or df.empty:
                continue

            # Detect header row by presence of required columns in first or second row.
            header_row_idx = None
            if len(df.index) >= 1 and _looks_like_transactions_header(df.iloc[0].tolist()):
                header_row_idx = 0
            elif len(df.index) >= 2 and _looks_like_transactions_header(df.iloc[1].tolist()):
                header_row_idx = 1

            if header_row_idx is None:
                continue  # not the Transactions table

            # Set header and drop header row
            df2 = df.copy()
            df2.columns = df2.iloc[header_row_idx]
            df2 = df2.iloc[header_row_idx + 1 :].reset_index(drop=True)

            # Standardize column names
            df2 = df2.rename(columns={c: _norm(c) for c in df2.columns})
            df2 = df2.rename(
                columns={
                    "date": "Date",
                    "description": "Description",
                    "amount": "Amount",
                    "balance": "Balance",
                }
            )

            # Keep only required columns in order
            keep = [c for c in ["Date", "Description", "Amount", "Balance"] if c in df2.columns]
            if set(keep) != {"Date", "Description", "Amount", "Balance"}:
                continue

            df2 = df2[keep]
            df2["source_file"] = pdf_file
            df2["table_index"] = i

            combined.append(df2)

    if not combined:
        raise RuntimeError(
            "No Transactions tables found. Check that the PDFs contain a table with headers: "
            "Date, Description, Amount, Balance."
        )

    return pd.concat(combined, ignore_index=True)


def save_to_excel(df: pd.DataFrame, output_path: str):
    df.to_excel(output_path, index=False, sheet_name="Transactions")


if __name__ == "__main__":
    args = parse_args()
    combined_df = extract_transactions_tables(args.input)
    save_to_excel(combined_df, args.output)
    print(f"Wrote {args.output} with {len(combined_df)} rows")
