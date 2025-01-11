name: Pages - TOP 10

on:
  schedule:
    - cron: "0 2 * * *"  # runs everyday at 2:00 AM UTC, corresponds to 9:00 PM ET during Standard Time
  workflow_dispatch:  # Allows manual triggering if needed

permissions:
  contents: write

jobs:
  update-html:
    runs-on: ubuntu-latest

    steps:
      - name: Checkout repository
        uses: actions/checkout@v3

      - name: Set up Python
        uses: actions/setup-python@v4
        with:
          python-version: '3.x'

      - name: Install dependencies
        run: pip install datetime openpyxl

      - name: Run Python script
        run: python scripts/pages-top10.py

      - name: Commit and push changes
        run: |
          git config --global user.name "GitHub Actions Bot"
          git config --global user.email "actions@github.com"
          git add ./top10/index.html
          git commit -m "Updated HTML with latest data"
          git push
