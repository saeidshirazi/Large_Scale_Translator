import pandas as pd
import asyncio
from deep_translator import GoogleTranslator
from rich.console import Console
from rich.progress import (
    Progress,
    SpinnerColumn,
    BarColumn,
    TextColumn,
    TimeElapsedColumn,
    TimeRemainingColumn
)
from rich.panel import Panel
from rich.table import Table
from rich import box
import json
import os
import time
import sys
import re

# Force UTF-8 encoding for Windows console
if sys.platform == "win32":
    sys.stdout.reconfigure(encoding='utf-8')

# =========================================================
# CONFIG
# =========================================================

INPUT_FILE = "unspsc.xlsx"
OUTPUT_FILE = "unspsc_persian.xlsx"

CACHE_FILE = "translation_cache.json"

# -------------------------------
# TEST MODE
# -------------------------------
TEST_MODE = False
TEST_ROWS = 100

# -------------------------------
# PERFORMANCE SETTINGS (MAX SPEED)
# -------------------------------
MAX_CHARS_PER_BATCH = 4500  # Safe but fast limit
CONCURRENT_TASKS = 8        # Optimized for high speed
MAX_RETRIES = 5             # Retry if Google blocks
BACKOFF_FACTOR = 2          # Exponential backoff

# Unique separator (Rare but stable)
SEPARATOR = " [[[ ]]] "

# =========================================================
# RICH CONSOLE
# =========================================================

console = Console()

console.print(
    Panel.fit(
        "[bold cyan]Excel Translator to Persian 🇮🇷[/bold cyan]\n"
        "[green]Fast • Cached • Async • Resumable[/green]",
        border_style="bright_blue"
    )
)

# =========================================================
# LOAD EXCEL
# =========================================================

console.print("[bold yellow]Loading Excel file...[/bold yellow]")

df = pd.read_excel(INPUT_FILE, dtype=str)

# =========================================================
# TEST MODE
# =========================================================

if TEST_MODE:
    df = df.head(TEST_ROWS)

    console.print(
        f"[bold magenta]TEST MODE ENABLED[/bold magenta] "
        f"→ Using first {TEST_ROWS} rows only"
    )
else:
    console.print(
        "[bold green]FULL DOCUMENT MODE ENABLED[/bold green]"
    )

# =========================================================
# LOAD CACHE
# =========================================================

if os.path.exists(CACHE_FILE):

    console.print(
        "[bold cyan]Loading translation cache...[/bold cyan]"
    )

    with open(CACHE_FILE, "r", encoding="utf-8") as f:
        translation_map = json.load(f)

else:
    translation_map = {}

console.print(
    f"[green]Cached translations:[/green] "
    f"{len(translation_map):,}"
)

# =========================================================
# EXTRACT TRANSLATABLE TEXTS
# =========================================================

console.print(
    "[bold yellow]Extracting translatable texts...[/bold yellow]"
)

all_values = pd.unique(df.values.ravel())

unique_texts = set()

for val in all_values:

    if pd.isna(val):
        continue

    text = str(val).strip()

    if not text:
        continue

    # Skip numbers
    if text.replace(".", "").isdigit():
        continue

    # Skip short text
    if len(text) < 3:
        continue

    # Skip URLs
    if text.startswith("http"):
        continue

    # Must contain alphabetic chars
    if not any(c.isalpha() for c in text):
        continue

    # Skip mostly numeric strings
    digit_ratio = sum(c.isdigit() for c in text) / len(text)

    if digit_ratio > 0.4:
        continue

    # Skip SMILES chemical formulas (common in this dataset)
    if "C" in text and "(" in text and ")" in text and any(c.isdigit() for c in text):
        if len(text) > 20 and not " " in text:
            continue

    unique_texts.add(text)

# Remove already cached
unique_texts = [
    text for text in unique_texts
    if text not in translation_map
]

# =========================================================
# INFO TABLE
# =========================================================

table = Table(
    title="Translation Statistics",
    box=box.ROUNDED
)

table.add_column("Metric", style="cyan")
table.add_column("Value", style="green")

table.add_row(
    "New Texts To Translate",
    f"{len(unique_texts):,}"
)

table.add_row(
    "Cached Translations",
    f"{len(translation_map):,}"
)

table.add_row(
    "Max Chars/Batch",
    str(MAX_CHARS_PER_BATCH)
)

table.add_row(
    "Concurrent Tasks",
    str(CONCURRENT_TASKS)
)

table.add_row(
    "Auto-Retry Logic",
    f"Enabled ({MAX_RETRIES} attempts)"
)

console.print(table)

# =========================================================
# CREATE BATCHES (SMART)
# =========================================================

def create_smart_batches(texts, max_chars):
    batches = []
    current_batch = []
    current_len = 0
    
    for text in texts:
        text_len = len(text) + len(SEPARATOR)
        
        # If a single text is too long, it must be in its own batch
        if text_len > max_chars:
            if current_batch:
                batches.append(current_batch)
                current_batch = []
                current_len = 0
            batches.append([text])
            continue

        if current_len + text_len > max_chars:
            batches.append(current_batch)
            current_batch = [text]
            current_len = text_len
        else:
            current_batch.append(text)
            current_len += text_len
            
    if current_batch:
        batches.append(current_batch)
    return batches

batches = create_smart_batches(unique_texts, MAX_CHARS_PER_BATCH)

console.print(
    f"[bold green]Total smart batches:[/bold green] "
    f"{len(batches):,}"
)

# =========================================================
# TRANSLATOR
# =========================================================

# =========================================================
# TRANSLATE SINGLE BATCH (WITH RETRY & BACKOFF)
# =========================================================

def sync_translate_batch(batch):
    if not batch:
        return {}

    # Per-task translator instance for better concurrency
    local_translator = GoogleTranslator(source='auto', target='fa')

    for attempt in range(MAX_RETRIES):
        try:
            # If single item
            if len(batch) == 1:
                return {batch[0]: local_translator.translate(batch[0])}

            joined_text = SEPARATOR.join(batch)
            translated_text = local_translator.translate(joined_text)
            
            if not translated_text:
                raise Exception("Empty result")

            parts = re.split(r'\s*\[\[\[\s*\]\]\]\s*', translated_text)
            parts = [p.strip() for p in parts if p.strip()]

            if len(parts) == len(batch):
                return dict(zip(batch, parts))
            
            # If slightly off, try simple split
            parts = translated_text.split(SEPARATOR.strip())
            parts = [p.strip() for p in parts if p.strip()]
            if len(parts) == len(batch):
                return dict(zip(batch, parts))

            # If still off, try one-by-one but fast
            raise Exception("Separator mismatch")

        except Exception as e:
            if attempt < MAX_RETRIES - 1:
                wait_time = BACKOFF_FACTOR ** attempt
                time.sleep(wait_time)
                continue
            
            # Final fallback: slow one-by-one
            result = {}
            for text in batch:
                try:
                    result[text] = local_translator.translate(text)
                except:
                    result[text] = text
            return result
    return {}

# =========================================================
# ASYNC WRAPPER
# =========================================================

async def async_translate_batch(batch, semaphore):

    async with semaphore:

        loop = asyncio.get_running_loop()

        result = await loop.run_in_executor(
            None,
            sync_translate_batch,
            batch
        )

        return result

# =========================================================
# MAIN TRANSLATION
# =========================================================

async def main():

    semaphore = asyncio.Semaphore(CONCURRENT_TASKS)

    tasks = [
        async_translate_batch(batch, semaphore)
        for batch in batches
    ]

    with Progress(
        SpinnerColumn(),
        TextColumn("[bold cyan]{task.description}"),
        BarColumn(bar_width=40),
        TextColumn("[progress.percentage]{task.percentage:>3.0f}%"),
        TextColumn("• [yellow]{task.completed}/{task.total} batches"),
        TimeElapsedColumn(),
        TimeRemainingColumn(),
        console=console
    ) as progress:

        task = progress.add_task(
            "🚀 Racing...",
            total=len(tasks)
        )

        completed = 0
        start_time = time.time()

        for coro in asyncio.as_completed(tasks):

            result = await coro
            translation_map.update(result)
            completed += 1
            progress.advance(task)

            # Update description with speed
            elapsed = time.time() - start_time
            if elapsed > 0:
                bpm = (completed / elapsed) * 60
                progress.update(task, description=f"🚀 Racing... [bold green]{bpm:.1f} bpm")

            # Save cache every 20 batches (optimized for speed)
            if completed % 20 == 0:
                with open(CACHE_FILE, "w", encoding="utf-8") as f:
                    json.dump(translation_map, f, ensure_ascii=False)

# =========================================================
# RUN TRANSLATION
# =========================================================

console.print(
    "\n[bold yellow]Starting translation...[/bold yellow]\n"
)

asyncio.run(main())

console.print(
    "\n[bold green]Translation completed![/bold green]"
)

# =========================================================
# APPLY TRANSLATIONS
# =========================================================

console.print(
    "[bold yellow]Applying translations...[/bold yellow]"
)

df = df.map(
    lambda x: translation_map.get(str(x).strip(), x)
    if pd.notna(x)
    else x
)

# =========================================================
# SAVE CACHE FINAL
# =========================================================

with open(CACHE_FILE, "w", encoding="utf-8") as f:
    json.dump(translation_map, f, ensure_ascii=False)

# =========================================================
# SAVE EXCEL
# =========================================================

console.print(
    "[bold yellow]Saving translated Excel file...[/bold yellow]"
)

df.to_excel(
    OUTPUT_FILE,
    index=False
)

# =========================================================
# DONE
# =========================================================

console.print(
    Panel.fit(
        f"[bold green]DONE![/bold green]\n\n"
        f"[cyan]Saved as:[/cyan]\n"
        f"{OUTPUT_FILE}",
        border_style="green"
    )
)