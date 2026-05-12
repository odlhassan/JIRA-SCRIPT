from pathlib import Path
import os

from report_server import create_report_server_app

base_dir = Path(__file__).resolve().parent
folder_raw = os.getenv("REPORT_HTML_DIR", "report_html")
app = create_report_server_app(base_dir=base_dir, folder_raw=folder_raw)