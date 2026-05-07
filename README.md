# Excel Translator to Persian 🇮🇷

Fast, resumable, async Excel (`.xlsx`) translator for huge datasets using Google Translate.

Built for:
- Massive Excel files
- Hundreds of thousands of rows
- Smart batching
- Automatic text detection
- Async parallel translation
- Translation caching
- Beautiful terminal UI with Rich

---

# Features

✅ Translate large `.xlsx` files into Persian  
✅ Async parallel translation engine  
✅ Smart batching based on character limits  
✅ Automatic detection of translatable text  
✅ Skips:
- numbers
- IDs
- URLs
- codes
- chemical formulas
- mostly numeric strings

✅ Persistent cache system (resume anytime)  
✅ Automatic retry with exponential backoff  
✅ Fancy interactive terminal UI using Rich  
✅ Progress bars with speed metrics  
✅ Test mode for quick validation  
✅ Works well with huge datasets  

---

# Demo

```text
╭────────────────────────────────────╮
│ Excel Translator to Persian 🇮🇷 │
│ Fast • Cached • Async • Resumable │
╰────────────────────────────────────╯

Loading Excel file...
FULL DOCUMENT MODE ENABLED

Translation Statistics
╭───────────────────────┬──────────╮
│ Metric                │ Value    │
├───────────────────────┼──────────┤
│ New Texts To Translate│ 299,736  │
│ Cached Translations   │ 120,443  │
│ Max Chars/Batch       │ 4500     │
│ Concurrent Tasks      │ 8        │
│ Auto-Retry Logic      │ Enabled  │
╰───────────────────────┴──────────╯

🚀 Racing... ━━━━━━━━━━━━━━━━ 52%
```

---

# Installation

Install dependencies:

```bash
pip install pandas openpyxl deep-translator rich
```

---

# Usage

Place your Excel file in the same directory as the script.

Example:

```text
project/
│
├── Translator.py
├── unspsc.xlsx
```

Run:

```bash
python Translator.py
```

Output:

```text
unspsc_persian.xlsx
```

---

# Configuration

Edit the configuration section inside the script:

```python
INPUT_FILE = "unspsc.xlsx"
OUTPUT_FILE = "unspsc_persian.xlsx"

CACHE_FILE = "translation_cache.json"
```

---

# Test Mode

Quickly test the translator on a few rows.

Enable:

```python
TEST_MODE = True
TEST_ROWS = 100
```

Disable for full translation:

```python
TEST_MODE = False
```

---

# Performance Settings

```python
MAX_CHARS_PER_BATCH = 4500
CONCURRENT_TASKS = 8
MAX_RETRIES = 5
BACKOFF_FACTOR = 2
```

## Recommended Values

| System/Internet | Concurrent Tasks |
|---|---|
| Slow Internet | 3–5 |
| Normal Internet | 5–8 |
| Fast Internet | 8–12 |

⚠ Very high concurrency may trigger Google rate limits.

---

# Smart Optimizations

The translator automatically:

## Detects Real Text

It skips:
- numbers
- short tokens
- URLs
- IDs
- chemical formulas
- mostly numeric strings

Example skipped values:

```text
12345
AB-991-X
https://example.com
C6H12O6
```

---

# Smart Batch Creation

Instead of fixed-size batches, the script:

- dynamically creates batches
- respects Google Translate limits
- maximizes throughput
- reduces failures

---

# Retry + Backoff System

If Google temporarily blocks requests:

- retries automatically
- uses exponential backoff
- falls back to one-by-one translation if needed

---

# Translation Cache

Translations are automatically saved into:

```text
translation_cache.json
```

Benefits:
- Resume interrupted translations
- Avoid retranslating duplicates
- Massive speed improvement
- Safe recovery after crashes

---

# Resume Support

You can stop the script anytime.

Next run will continue automatically using the cache.

---

# Performance

Typical improvements:

| Optimization | Effect |
|---|---|
| Async Translation | Huge speedup |
| Smart Batching | Fewer requests |
| Translation Cache | Avoid duplicate work |
| Auto Text Detection | Reduces translation count |
| Retry System | Better stability |

---


# Project Structure

```text
project/
│
├── Translator.py
```

---

# Example Workflow

1. Load Excel
2. Detect valid text automatically
3. Remove already cached translations
4. Create smart batches
5. Translate asynchronously
6. Save cache progressively
7. Apply translations
8. Export Persian Excel file

---

# Requirements

- Python 3.9+
- Internet connection

---

# License

MIT License

---

# Author

Built for fast large-scale Excel translation workflows. Made with Love for Persian Community!
