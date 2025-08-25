# test_openai_connect.py
from openai import OpenAI

API_KEY = "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"

try:
    client = OpenAI(api_key=API_KEY)
    # Easiest lightweight call that proves auth + network
    models = client.models.list()
    print("OK ✅ Models count:", len(models.data))
except Exception as e:
    print("FAIL ❌", repr(e))
