import os
import json
import datetime
from typing import List, Dict, Any
import requests
import feedparser
from bs4 import BeautifulSoup
from docx import Document
from docx.shared import Pt, Inches, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from google.oauth2 import service_account
from googleapiclient.discovery import build
from googleapiclient.http import MediaFileUpload
from google import genai
from google.genai import types

# Targeted RSS and site news feeds for Pakistan business sources
RSS_FEEDS = [
    {"source": "Business Recorder", "url": "https://www.brecorder.com/feeds/latest-news/"},
    {"source": "Dawn Business", "url": "https://www.dawn.com/feeds/business/"},
    {"source": "The News International - Business", "url": "https://www.thenews.com.pk/rss/1/2"}, # Business section
    {"source": "Tribune Business", "url": "https://tribune.com.pk/feed/business"},
    {"source": "Pakistan Today - Profit", "url": "https://profit.pakistantoday.com.pk/feed/"},
]

def fetch_recent_feed_items(hours_window: int = 24) -> List[Dict[str, Any]]:
    """Fetch recent feed items within the last 24 hours."""
    articles = []
    cutoff_time = datetime.datetime.now(datetime.timezone.utc) - datetime.timedelta(hours=hours_window)
    
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64) AppleWebKit/537.36 (KHTML, like Gecko) Chrome/120.0.0.0 Safari/537.36"
    }

    for feed_info in RSS_FEEDS:
        try:
            resp = requests.get(feed_info["url"], headers=headers, timeout=12)
            if resp.status_code == 200:
                parsed = feedparser.parse(resp.content)
                for entry in parsed.entries[:25]:
                    published_dt = None
                    if hasattr(entry, 'published_parsed') and entry.published_parsed:
                        published_dt = datetime.datetime(*entry.published_parsed[:6], tzinfo=datetime.timezone.utc)
                    elif hasattr(entry, 'updated_parsed') and entry.updated_parsed:
                        published_dt = datetime.datetime(*entry.updated_parsed[:6], tzinfo=datetime.timezone.utc)
                    
                    if not published_dt or published_dt >= cutoff_time:
                        title = getattr(entry, 'title', '').strip()
                        link = getattr(entry, 'link', '').strip()
                        summary = getattr(entry, 'summary', '').strip()
                        
                        # Strip raw html from summary if present
                        if summary:
                            summary = BeautifulSoup(summary, "html.parser").get_text()

                        articles.append({
                            "source": feed_info["source"],
                            "title": title,
                            "link": link,
                            "published": published_dt.strftime("%Y-%m-%d %H:%M UTC") if published_dt else "Last 24 Hours",
                            "raw_summary": summary[:400]
                        })
        except Exception as e:
            print(f"[WARN] Error fetching {feed_info['source']}: {e}")

    return articles

def analyze_and_structure_news(articles: List[Dict[str, Any]], api_key: str) -> Dict[str, Any]:
    """Uses Gemini API to curate, summarize, assess market impact, and categorize news."""
    client = genai.Client(api_key=api_key)
    
    prompt = f"""
You are an expert investment analyst for equity investments in Pakistan focusing on:
economy, pharmaceutical, automobile, sugar, technology, food, and modaraba sectors.

Filter and analyze the following collected news items published in the last 24 hours.
Discard:
- Items older than 24 hours
- Non-material international/global news unless direct material impact on Pakistan (e.g. IMF, crude oil import bill)
- Fluff or irrelevant political drama with zero economic/business bearing

Selected sectors: Economy, Pharmaceutical, Automobile, Sugar, Technology, Food, Modaraba.

Segregate the items into strictly four categories:
1. "Main News" (Most critical macro/national business headlines)
2. "Economic News" (Monetary policy, fiscal updates, taxation, IMF, inflation, trade, currency)
3. "Sector-Specific News" (Specific developments in Pharma, Auto, Sugar, Tech, Food, Modaraba. IF NO news in these sectors, return an empty array [])
4. "Other Notables" (Other notable business/industry news if any)

For EACH item, provide:
- "headline": Concise professional headline
- "summary": 2 to 3 sentences summarizing the development
- "impact": 2 to 3 sentences detailing the specific impact on Pakistani markets, economy, or sector companies
- "source": Source publication name
- "publish_time": Publish date/time from the input
- "url": Source URL

Return output strictly in JSON format with keys:
"main_news", "economic_news", "sector_specific_news", "other_notables", "sources".
Where "sources" is a list of all referenced items with title and url.

Collected Articles Raw List:
{json.dumps(articles[:40], indent=2)}
"""

    response = None
    import time

    # Query available models or use prioritized list
    available_models = ["gemini-2.5-flash", "gemini-2.5-pro", "gemini-3.8-flash", "gemini-flash", "gemini-pro"]
    try:
        remote_models = [m.name.replace("models/", "") for m in client.models.list() if "generateContent" in (getattr(m, "supported_generation_methods", []) or getattr(m, "supported_actions", []))]
        if remote_models:
            # prioritize flash models first
            flash_models = [m for m in remote_models if "flash" in m]
            other_models = [m for m in remote_models if "flash" not in m]
            available_models = flash_models + other_models
            print(f"[*] Available remote models: {available_models[:6]}")
    except Exception as e:
        print(f"[*] Could not query model list dynamically ({e}), using default candidates.")

    for mod in available_models:
        for attempt in range(4):
            try:
                print(f"[*] Trying model {mod} (attempt {attempt+1})...")
                response = client.models.generate_content(
                    model=mod,
                    contents=prompt,
                    config=types.GenerateContentConfig(
                        response_mime_type="application/json"
                    )
                )
                if response and response.text:
                    print(f"[+] Successfully generated content using {mod}!")
                    break
            except Exception as err:
                err_str = str(err)
                print(f"[!] Error with {mod}: {err_str[:120]}...")
                if "404" in err_str:
                    # Model doesn't exist, skip immediately to next model
                    break
                # If 503 high demand or 429, wait with exponential backoff
                wait_time = (attempt + 1) * 6
                print(f"    Waiting {wait_time}s before retry...")
                time.sleep(wait_time)
        if response and response.text:
            break

    if not response or not response.text:
        raise RuntimeError("Failed to generate content from Gemini API across fallback models.")
    
    return json.loads(response.text)

def create_docx_report(data: Dict[str, Any], output_path: str):
    """Generates the structured Word (.docx) document according to specifications."""
    doc = Document()
    
    # Page setup
    sections = doc.sections
    for s in sections:
        s.top_margin = Inches(1)
        s.bottom_margin = Inches(1)
        s.left_margin = Inches(1)
        s.right_margin = Inches(1)

    today_str = datetime.datetime.now().strftime("%B %d, %Y")

    def format_run(run, font_name="Times New Roman", size_pt=12, bold=False, color_rgb=(0,0,0)):
        run.font.name = font_name
        run.font.size = Pt(size_pt)
        run.bold = bold
        run.font.color.rgb = RGBColor(*color_rgb)

    def add_article_block(item: Dict[str, Any]):
        p_head = doc.add_paragraph()
        p_head.paragraph_format.space_before = Pt(8)
        p_head.paragraph_format.space_after = Pt(2)
        r_head = p_head.add_run(item.get("headline", ""))
        format_run(r_head, bold=True, size_pt=12)

        p_meta = doc.add_paragraph()
        p_meta.paragraph_format.space_before = Pt(0)
        p_meta.paragraph_format.space_after = Pt(4)
        r_meta = p_meta.add_run(f"Source: {item.get('source', '')} | Published: {item.get('publish_time', '')}")
        format_run(r_meta, size_pt=10, color_rgb=(100, 100, 100))

        p_sum = doc.add_paragraph()
        p_sum.paragraph_format.space_before = Pt(2)
        p_sum.paragraph_format.space_after = Pt(3)
        r_sum_label = p_sum.add_run("Summary: ")
        format_run(r_sum_label, bold=True, size_pt=12)
        r_sum_val = p_sum.add_run(item.get("summary", ""))
        format_run(r_sum_val, size_pt=12)

        p_imp = doc.add_paragraph()
        p_imp.paragraph_format.space_before = Pt(2)
        p_imp.paragraph_format.space_after = Pt(10)
        r_imp_label = p_imp.add_run("Impact: ")
        format_run(r_imp_label, bold=True, size_pt=12)
        r_imp_val = p_imp.add_run(item.get("impact", ""))
        format_run(r_imp_val, size_pt=12)

    # 1. Main Heading & Date
    p_title = doc.add_paragraph()
    p_title.paragraph_format.space_before = Pt(0)
    p_title.paragraph_format.space_after = Pt(2)
    p_title.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_title = p_title.add_run("The Daily News Coverage")
    format_run(r_title, bold=True, size_pt=18)

    p_date = doc.add_paragraph()
    p_date.paragraph_format.space_before = Pt(0)
    p_date.paragraph_format.space_after = Pt(18)
    p_date.alignment = WD_ALIGN_PARAGRAPH.CENTER
    r_date = p_date.add_run(today_str)
    format_run(r_date, size_pt=12, color_rgb=(80, 80, 80))

    # Page 1: Main News
    p_sec1 = doc.add_paragraph()
    p_sec1.paragraph_format.space_after = Pt(6)
    r_sec1 = p_sec1.add_run("Main News")
    format_run(r_sec1, bold=True, size_pt=14)
    for it in data.get("main_news", []):
        add_article_block(it)

    # Page 2: Economic News
    doc.add_page_break()
    p_sec2 = doc.add_paragraph()
    p_sec2.paragraph_format.space_after = Pt(6)
    r_sec2 = p_sec2.add_run("Economic News")
    format_run(r_sec2, bold=True, size_pt=14)
    for it in data.get("economic_news", []):
        add_article_block(it)

    # Page 3: Sector-Specific News (Only if available)
    sector_news = data.get("sector_specific_news", [])
    if sector_news and len(sector_news) > 0:
        doc.add_page_break()
        p_sec3 = doc.add_paragraph()
        p_sec3.paragraph_format.space_after = Pt(6)
        r_sec3 = p_sec3.add_run("Sector-Specific News")
        format_run(r_sec3, bold=True, size_pt=14)
        for it in sector_news:
            add_article_block(it)

    # Page 4: Other Notables (If any)
    other_news = data.get("other_notables", [])
    if other_news and len(other_news) > 0:
        doc.add_page_break()
        p_sec4 = doc.add_paragraph()
        p_sec4.paragraph_format.space_after = Pt(6)
        r_sec4 = p_sec4.add_run("Other Notables")
        format_run(r_sec4, bold=True, size_pt=14)
        for it in other_news:
            add_article_block(it)

    # Final Page: Source List with links
    doc.add_page_break()
    p_src = doc.add_paragraph()
    p_src.paragraph_format.space_after = Pt(8)
    r_src = p_src.add_run("Sources & References")
    format_run(r_src, bold=True, size_pt=14)

    sources = data.get("sources", [])
    if not sources:
        # Fallback collect from all items
        for key in ["main_news", "economic_news", "sector_specific_news", "other_notables"]:
            for item in data.get(key, []):
                if item.get("url"):
                    sources.append({"title": item.get("headline", item.get("source", "")), "url": item.get("url")})

    for s in sources:
        p_link = doc.add_paragraph(style='List Bullet')
        p_link.paragraph_format.space_after = Pt(4)
        r_st = p_link.add_run(f"{s.get('title', 'Article')}: ")
        format_run(r_st, bold=True, size_pt=11)
        r_su = p_link.add_run(s.get('url', ''))
        format_run(r_su, size_pt=11, color_rgb=(0, 0, 180))

    doc.save(output_path)
    return output_path

def upload_to_google_drive(file_path: str, sa_json_str: str, folder_name: str = "Daily News") -> str:
    """Uploads file to Google Drive using a Service Account and returns the shareable web link."""
    sa_info = json.loads(sa_json_str)
    creds = service_account.Credentials.from_service_account_info(
        sa_info,
        scopes=["https://www.googleapis.com/auth/drive"]
    )
    service = build("drive", "v3", credentials=creds)

    folder_id = os.environ.get("GDRIVE_FOLDER_ID")
    
    if not folder_id:
        # Search for folder in accessible files/shared folders
        query = f"name = '{folder_name}' and mimeType = 'application/vnd.google-apps.folder' and trashed = false"
        res = service.files().list(
            q=query,
            spaces='drive',
            fields='files(id, name)',
            supportsAllDrives=True,
            includeItemsFromAllDrives=True
        ).execute()
        folders = res.get('files', [])
        if folders:
            folder_id = folders[0]['id']
        else:
            folder_metadata = {
                'name': folder_name,
                'mimeType': 'application/vnd.google-apps.folder'
            }
            folder = service.files().create(
                body=folder_metadata,
                fields='id',
                supportsAllDrives=True
            ).execute()
            folder_id = folder.get('id')

    # Upload file directly into the shared folder
    file_name = os.path.basename(file_path)
    file_metadata = {
        'name': file_name,
        'parents': [folder_id]
    }
    media = MediaFileUpload(
        file_path,
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        resumable=True
    )
    uploaded = service.files().create(
        body=file_metadata,
        media_body=media,
        fields='id, webViewLink',
        supportsAllDrives=True
    ).execute()

    return uploaded.get('webViewLink', f"https://drive.google.com/file/d/{uploaded.get('id')}/view")

def main():
    gemini_api_key = os.environ.get("GEMINI_API_KEY")
    gdrive_sa_json = os.environ.get("GDRIVE_SERVICE_ACCOUNT_JSON")

    if not gemini_api_key:
        raise ValueError("Missing GEMINI_API_KEY environment variable")
    if not gdrive_sa_json:
        raise ValueError("Missing GDRIVE_SERVICE_ACCOUNT_JSON environment variable")

    print("[*] 1. Gathering latest 24h news from Pakistan sources...")
    articles = fetch_recent_feed_items(hours_window=24)
    print(f"    Collected {len(articles)} candidate stories.")

    print("[*] 2. Analyzing & filtering via Gemini AI...")
    structured_data = analyze_and_structure_news(articles, gemini_api_key)

    today_formatted = datetime.datetime.now().strftime("%B %d, %Y")
    file_name = f"The Daily News Coverage - {today_formatted}.docx"
    output_docx_path = os.path.join(os.getcwd(), file_name)

    print(f"[*] 3. Compiling Word Document: {file_name}...")
    create_docx_report(structured_data, output_docx_path)

    print("[*] 4. Uploading to Google Drive ('Daily News' folder)...")
    drive_link = upload_to_google_drive(output_docx_path, gdrive_sa_json, folder_name="Daily News")
    print(f"[SUCCESS] Report successfully generated and uploaded!")
    print(f"Google Drive Link: {drive_link}")

if __name__ == "__main__":
    main()
