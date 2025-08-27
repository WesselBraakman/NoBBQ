#!/usr/bin/env python3
"""
excel_prompts_to_answers_gemini.py

- Reads prompts from Column A in an Excel sheet
- Sends each prompt to Google Gemini (1.5-flash by default)
- Writes the model's answer into Column B
- Skips rows where Column B already has an answer (resume)
- Saves results to a new file with "_with_answers.xlsx" suffix
- Prints progress while running

Prereqs:
  pip install google-generativeai pandas openpyxl
Env:
  set GOOGLE_API_KEY=your_key_here   (Windows PowerShell: $env:GOOGLE_API_KEY="...")
"""

import argparse
import os
import time
import pandas as pd
import google.generativeai as genai

DEFAULT_MODEL = "gemini-1.5-flash"   # or "gemini-1.5-pro" for higher quality

def init_client():
    api_key = os.getenv("GOOGLE_API_KEY")
    if not api_key:
        raise SystemExit("Missing GOOGLE_API_KEY environment variable.")
    genai.configure(api_key=api_key)

def call_gemini(prompt: str, model_name: str) -> str:
    """Send one prompt to Gemini and return the answer text (stateless per call)."""
    model = genai.GenerativeModel(model_name)
    resp = model.generate_content(prompt)
    # The SDK raises for blocked/safety in .prompt_feedback; also check candidates
    if hasattr(resp, "text") and resp.text:
        return resp.text.strip()
    # Fallback: concatenate candidate text parts if available
    parts = []
    for c in getattr(resp, "candidates", []) or []:
        for p in getattr(c, "content", {}).get("parts", []) or []:
            t = getattr(p, "text", "")
            if t:
                parts.append(t)
    return "".join(parts).strip()

def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("excel_path", help="Path to the Excel file")
    parser.add_argument("--sheet", default=0, help="Sheet name or index (default: first)")
    parser.add_argument("--model", default=DEFAULT_MODEL, help="Gemini model (default: gemini-1.5-flash)")
    parser.add_argument("--sleep", type=float, default=0.3, help="Pause (seconds) between calls")
    args = parser.parse_args()

    init_client()

    # Load Excel
    df = pd.read_excel(args.excel_path, sheet_name=args.sheet, engine="openpyxl")
    if df.shape[1] < 1:
        raise SystemExit("No column A found.")
    if df.shape[1] < 2:
        df.insert(1, "Answer", "")

    total = len(df)
    processed = 0

    for idx, val in df.iloc[:, 0].items():
        # Count progress regardless, so the counter matches rows in the sheet
        try:
            if pd.isna(val) or not str(val).strip():
                processed += 1
                print(f"Processed {processed}/{total} rows... (skipped empty A)")
                continue
            if not pd.isna(df.iat[idx, 1]) and str(df.iat[idx, 1]).strip():
                processed += 1
                print(f"Processed {processed}/{total} rows... (resume skip)")
                continue

            prompt = str(val).strip()
            try:
                answer = call_gemini(prompt, args.model)
                df.iat[idx, 1] = answer if answer else "(empty response)"
            except Exception as e:
                df.iat[idx, 1] = f"ERROR: {e}"

            processed += 1
            print(f"Processed {processed}/{total} rows...")
            time.sleep(args.sleep)
        except KeyboardInterrupt:
            print("Interrupted by user. Writing partial results...")
            break

    # Save to a new file
    out_path = args.excel_path.replace(".xlsx", "_with_answers.xlsx")
    with pd.ExcelWriter(out_path, engine="openpyxl", mode="w") as writer:
        df.to_excel(
            writer,
            sheet_name=args.sheet if isinstance(args.sheet, str) else writer.book.sheetnames[0],
            index=False,
        )

    print(f"✅ Done. Answers written to {out_path}")

if __name__ == "__main__":
    main()
