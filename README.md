

# Excel Translator to Persian 🇮🇷

A fast and optimized Python tool for translating large Excel (`.xlsx`) files into Persian using Google Translate.

This project is designed for:
- Huge Excel files
- Hundreds of thousands of rows
- Fast parallel translation
- Progress tracking
- Batch processing
- Low memory overhead

---

# Features

✅ Translate entire `.xlsx` files to Persian  
✅ Optimized for very large datasets  
✅ Parallel translation using asyncio + threading  
✅ Batch translation for huge speed improvements  
✅ Progress bar with `tqdm`  
✅ Test mode for quick validation  
✅ Preserves Excel structure and columns  
✅ Automatically skips duplicate translations  

---

# Installation

Install dependencies:

```bash
pip install pandas openpyxl deep-translator tqdm
```

---

# Usage

Place your Excel file in the same directory.

Example:

```text
unspsc.xlsx
translator.py
```

Run:

```bash
python translator.py
```

---

# Configuration

Inside the script:

```python
INPUT_FILE = "unspsc.xlsx"
OUTPUT_FILE = "unspsc_persian.xlsx"
```

---

# Test Mode

Translate only first few rows for testing:

```python
TEST_MODE = True
TEST_ROWS = 10
```

For full translation:

```python
TEST_MODE = False
```

---

# Performance Settings

You can tune performance here:

```python
BATCH_SIZE = 20
CONCURRENT_TASKS = 20
```

Recommended values:

| Internet Speed | BATCH_SIZE | CONCURRENT_TASKS |
|---|---|---|
| Slow | 10 | 10 |
| Medium | 20 | 20 |
| Fast | 30 | 30 |

⚠ Very high concurrency may trigger Google rate limiting.

---

# How It Works

The tool:

1. Loads the Excel file
2. Extracts all unique text values
3. Translates texts in batches
4. Uses async concurrency for speed
5. Applies translations back to the dataframe
6. Saves translated `.xlsx`

This dramatically reduces translation requests.

Example:

Instead of:

```text
300,000 API requests
```

It may only send:

```text
15,000 batch requests
```

---

# Example Output

Input:

| Product |
|---|
| Computer |
| Keyboard |

Output:

| Product |
|---|
| کامپیوتر |
| صفحه کلید |

---

# Progress Bar

Example:

```text
Translating: 45%|█████████████ | 450/1000
```

---

# Limitations

- Uses unofficial Google Translate access
- Very large datasets may still take time
- Google may temporarily rate limit requests
- Formatting/styles are not preserved perfectly

---

# Recommended Improvements

For even better performance:

- Convert `.xlsx` → `.csv`
- Use paid APIs like:
  - DeepL
  - OpenAI
  - Google Cloud Translation API

---

# License

MIT License

---

# Author

Created for fast large-scale Excel translation workflows. Made by Love for Persian Community!
