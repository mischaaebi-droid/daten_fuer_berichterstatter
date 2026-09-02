         xxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxxx

          # Einmaliger historischer Import:
# 1. Januar 2026 bis einschliesslich 1. September 2026

name: Bundesgericht Backfill 2026

on:
  workflow_dispatch:

permissions:
  contents: write

jobs:
  scrape:
    runs-on: ubuntu-latest
    timeout-minutes: 360

    steps:
      - name: Repository laden
        uses: actions/checkout@v4

      - name: Python installieren
        uses: actions/setup-python@v5
        with:
          python-version: '3.13'

      - name: Python-Pakete installieren
        run: |
          pip install playwright beautifulsoup4 lxml pandas openpyxl
          playwright install chromium

      - name: Alle Tage vom 1. Januar bis 1. September 2026 scrapen
        shell: bash
        run: |
          CURRENT_DATE="2026-01-01"
          END_DATE="2026-09-01"

          while [[ "$CURRENT_DATE" < "$END_DATE" || "$CURRENT_DATE" == "$END_DATE" ]]; do
            SCRIPT_DATE=$(date -d "$CURRENT_DATE" +"%Y%m%d")

            echo "========================================"
            echo "Scrape Datum: $CURRENT_DATE"
            echo "========================================"

            python bundesgericht_tagesurteile.py --date "$SCRIPT_DATE"

            CURRENT_DATE=$(date -I -d "$CURRENT_DATE + 1 day")
          done

      - name: JSON auf GitHub speichern
        run: |
          git config user.name "github-actions[bot]"
          git config user.email "41898282+github-actions[bot]@users.noreply.github.com"
          git add data/decisions-*.json
          git commit -m "Bundesgerichtsdaten Januar bis September 2026 ergänzen" || exit 0
          git push

