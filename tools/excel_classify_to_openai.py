#!/usr/bin/env python3
"""
excel_classify_to_openai.py

Reads column A from Excel, sends each row's text to OpenAI,
and stores the model's response in column B.
"""

import argparse
import re
import time
from typing import List

import pandas as pd
from tenacity import retry, stop_after_attempt, wait_exponential, retry_if_exception_type

from openai import OpenAI

# ---------------------------
# 🔑 Hardcode your API key here
OPENAI_API_KEY = "apikey"
# ---------------------------

client = OpenAI(api_key=OPENAI_API_KEY)

LABEL_PATTERN = re.compile(r"\bans\d+\b", flags=re.IGNORECASE)


def extract_labels_from_text(text: str) -> List[str]:
    labels = LABEL_PATTERN.findall(text or "")
    seen = set()
    ordered = []
    for lab in labels:
        lab = lab.lower()
        if lab not in seen:
            seen.add(lab)
            ordered.append(lab)
    if not ordered:
        ordered = ["ans0", "ans1", "ans2"]
    return ordered


def build_prompt(row_text: str, allowed_labels: List[str]) -> str:
    return (
        "Du er en streng klassifikator. "
        "Returner KUN én etikett, nøyaktig som skrevet, uten forklaring, uten ekstra tegn.\n\n"
        f"Velg hvilken av disse etikettene som passer best til teksten: {', '.join(allowed_labels)}.\n\n"
        f"Tekst:\n{row_text}\n\n"
        f"Returner kun én: {', '.join(allowed_labels)}"
    )


class TransientOpenAIError(Exception):
    pass


@retry(
    reraise=True,
    stop=stop_after_attempt(5),
    wait=wait_exponential(multiplier=1, min=1, max=20),
    retry=retry_if_exception_type(TransientOpenAIError),
)
def call_openai_strict_label(prompt: str, model: str) -> str:
    try:
        resp = client.responses.create(
            model=model,
            input=prompt,   # ✅ use input instead of messages
            temperature=0,
        )

        output_text = getattr(resp, "output_text", None)
        if not output_text:
            pieces = []
            for item in getattr(resp, "output", []) or []:
                for c in getattr(item, "content", []) or []:
                    if getattr(c, "type", "") in {"output_text", "text"} and getattr(c, "text", ""):
                        pieces.append(c.text)
            output_text = "".join(pieces).strip()

        if not output_text:
            raise ValueError("Empty response from model")

        return output_text.strip()

    except Exception as e:
        msg = str(e).lower()
        if any(term in msg for term in ["rate limit", "timeout", "temporar", "try again", "overloaded", "502", "503", "504"]):
            raise TransientOpenAIError(e)
        raise


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument("excel_path", help="Path to the Excel file (e.g., data.xlsx)")
    parser.add_argument("--sheet", default=0, help="Sheet name or index (default: first sheet)")
    parser.add_argument("--model", default="gpt-5", help="OpenAI model (default: gpt-5)")
    args = parser.parse_args()

    df = pd.read_excel(args.excel_path, sheet_name=args.sheet, engine="openpyxl")

    if df.shape[1] < 1:
        raise SystemExit("No column A found in the sheet.")

    if df.shape[1] < 2:
        df.insert(1, "ModelSvar", "")

    for idx, val in df.iloc[:, 0].items():
        if pd.isna(val) or not str(val).strip():
            continue

        row_text = str(val).strip()
        allowed = extract_labels_from_text(row_text)
        prompt = build_prompt(row_text, allowed)

        try:
            label = call_openai_strict_label(prompt, model=args.model)
            m = LABEL_PATTERN.search(label)
            clean_label = m.group(0).lower() if m else label.strip()
            df.iat[idx, 1] = clean_label
        except Exception as e:
            df.iat[idx, 1] = f"ERROR: {e}"
            time.sleep(1)

    out_path = args.excel_path.replace(".xlsx", "_with_results.xlsx")
    with pd.ExcelWriter(out_path, engine="openpyxl", mode="w") as writer:
        df.to_excel(writer, sheet_name=args.sheet if isinstance(args.sheet, str) else writer.book.sheetnames[0], index=False)

    print(f"✅ Done. Results written to {out_path}")


if __name__ == "__main__":
    main()
