# Q2 2025 Machine Learning Map

## Easiest way to open the photo gallery

Double-click `open_gallery.bat`.

It will:

1. Read `Instemming Huize G (Responses).xlsx`
2. Make `gallery.html`
3. Open `gallery.html` in your browser

If it says pandas or openpyxl is missing, run this once:

```powershell
pip install pandas openpyxl
```

Then double-click `open_gallery.bat` again.

## Manual option

You can also run:

```powershell
python make_gallery.py
```

Then open `gallery.html`.

The script keeps working even when a Google Drive image cannot be downloaded. In that case it shows a button that opens the original photo link.
