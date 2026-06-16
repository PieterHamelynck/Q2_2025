from __future__ import annotations

import base64
import html
import re
import sys
from pathlib import Path
from urllib.error import HTTPError, URLError
from urllib.parse import parse_qs, urlparse
from urllib.request import Request, urlopen


EXCEL_FILE = Path("Instemming Huize G (Responses).xlsx")
OUTPUT_FILE = Path("gallery.html")
INTRO_COLUMN = "Vertel hier wat leuks over jezelf"
PHOTO_COLUMN = "Fotootjes van jou!"


def extract_drive_file_id(url: str) -> str | None:
    parsed = urlparse(url)
    if "drive.google.com" not in parsed.netloc:
        return None

    if parsed.path == "/open":
        return parse_qs(parsed.query).get("id", [None])[0]

    path_parts = [part for part in parsed.path.split("/") if part]
    if len(path_parts) >= 3 and path_parts[0] == "file" and path_parts[1] == "d":
        return path_parts[2]

    if parsed.path == "/uc":
        return parse_qs(parsed.query).get("id", [None])[0]

    match = re.search(r"/d/([a-zA-Z0-9_-]+)", parsed.path)
    if match:
        return match.group(1)

    return None


def candidate_image_urls(url: str) -> list[str]:
    file_id = extract_drive_file_id(url)
    if not file_id:
        return [url]

    return [
        f"https://drive.google.com/thumbnail?id={file_id}&sz=w1200",
        f"https://drive.google.com/uc?export=view&id={file_id}",
        f"https://drive.google.com/uc?export=download&id={file_id}",
        url,
    ]


def detect_image_type(data: bytes, content_type: str) -> str | None:
    content_type = content_type.lower()
    if content_type.startswith("image/"):
        return content_type.split(";", 1)[0]
    if data[:4] == b"\x89PNG":
        return "image/png"
    if data[:3] == b"\xff\xd8\xff":
        return "image/jpeg"
    if data[:6] in (b"GIF87a", b"GIF89a"):
        return "image/gif"
    if data[:12] == b"RIFF" and data[8:12] == b"WEBP":
        return "image/webp"
    return None


def fetch_image_as_data_uri(url: str) -> str | None:
    headers = {
        "User-Agent": "Mozilla/5.0",
        "Accept": "image/avif,image/webp,image/apng,image/*,*/*;q=0.8",
    }

    for candidate in candidate_image_urls(url):
        try:
            request = Request(candidate, headers=headers)
            with urlopen(request, timeout=15) as response:
                content_type = response.headers.get("Content-Type") or ""
                data = response.read()

            image_type = detect_image_type(data, content_type)
            if not data or image_type is None:
                continue

            encoded = base64.b64encode(data).decode("ascii")
            return f"data:{image_type};base64,{encoded}"
        except (HTTPError, URLError, TimeoutError, ValueError):
            continue

    return None


def photo_links(value: object) -> list[str]:
    if value is None:
        return []
    text = str(value).strip()
    if not text or text.lower() == "nan":
        return []
    return [url.strip() for url in text.split(",") if url.strip()]


def build_card(row_number: int, intro: object, photos: object) -> str:
    intro_text = html.escape(str(intro).strip())
    image_tags = []

    for photo_number, url in enumerate(photo_links(photos), start=1):
        safe_url = html.escape(url, quote=True)
        data_uri = fetch_image_as_data_uri(url)
        if data_uri:
            image_tags.append(
                f'<a href="{safe_url}" target="_blank" rel="noopener noreferrer">'
                f'<img src="{data_uri}" alt="Foto {photo_number}" class="photo-item" loading="lazy">'
                "</a>"
            )
        else:
            image_tags.append(
                f'<a href="{safe_url}" target="_blank" rel="noopener noreferrer" '
                f'class="photo-fallback">Foto {photo_number} openen</a>'
            )

    images_html = "".join(image_tags) or "<em>Geen foto beschikbaar</em>"
    return f"""
        <article class="person-card">
            <h2>Persoon {row_number}</h2>
            <p>{intro_text}</p>
            <div class="photo-grid">{images_html}</div>
        </article>
    """


def page_html(cards: list[str]) -> str:
    return f"""<!doctype html>
<html lang="nl">
<head>
<meta charset="utf-8">
<meta name="viewport" content="width=device-width, initial-scale=1">
<title>Fotootjes + Intro Tekst</title>
<style>
    body {{
        margin: 0;
        background: #f4f6f8;
        color: #1f2933;
        font-family: Arial, sans-serif;
    }}
    main {{
        max-width: 1040px;
        margin: 0 auto;
        padding: 24px 16px 40px;
    }}
    h1 {{
        margin: 0 0 18px;
        font-size: 28px;
        font-weight: 700;
    }}
    .person-card {{
        margin-bottom: 16px;
        padding: 16px;
        border: 1px solid #d7dee8;
        border-radius: 8px;
        background: #ffffff;
    }}
    .person-card h2 {{
        margin: 0 0 8px;
        font-size: 18px;
    }}
    .person-card p {{
        margin: 0 0 12px;
        line-height: 1.5;
        white-space: pre-wrap;
        overflow-wrap: anywhere;
    }}
    .photo-grid {{
        display: flex;
        flex-wrap: wrap;
        gap: 10px;
    }}
    .photo-item,
    .photo-fallback {{
        width: 150px;
        height: 150px;
        border: 1px solid #cbd5e1;
        border-radius: 8px;
        background: #eef2f7;
    }}
    .photo-item {{
        display: block;
        object-fit: cover;
    }}
    .photo-fallback {{
        display: inline-flex;
        align-items: center;
        justify-content: center;
        box-sizing: border-box;
        padding: 10px;
        color: #1d4ed8;
        text-align: center;
        text-decoration: none;
    }}
</style>
</head>
<body>
<main>
    <h1>Fotootjes + Intro Tekst</h1>
    {"".join(cards)}
</main>
</body>
</html>
"""


def main() -> int:
    try:
        import pandas as pd
    except ImportError:
        print("Pandas is missing. Run: pip install pandas openpyxl")
        return 1

    if not EXCEL_FILE.exists():
        print(f"Could not find {EXCEL_FILE}. Keep this script next to the Excel file.")
        return 1

    print(f"Reading {EXCEL_FILE}...")
    dataframe = pd.read_excel(EXCEL_FILE)
    missing_columns = [col for col in (INTRO_COLUMN, PHOTO_COLUMN) if col not in dataframe.columns]
    if missing_columns:
        print("Missing expected column(s): " + ", ".join(missing_columns))
        return 1

    dataframe = dataframe.dropna(subset=[INTRO_COLUMN])
    cards = []
    for row_number, row in dataframe.iterrows():
        print(f"Making card {len(cards) + 1} of {len(dataframe)}...")
        cards.append(build_card(row_number, row[INTRO_COLUMN], row.get(PHOTO_COLUMN)))

    OUTPUT_FILE.write_text(page_html(cards), encoding="utf-8")
    print(f"Done. Open {OUTPUT_FILE.resolve()} in your browser.")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())
