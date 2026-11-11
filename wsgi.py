from pathlib import Path
import os

from maintenance_service import MaintenanceGateway

base_dir = Path(__file__).resolve().parent
folder_raw = os.getenv("REPORT_HTML_DIR", "report_html")

def build_app():
    # Serve the maintenance shell while the reporting app imports and initializes.
    from report_server import create_report_server_app
    return create_report_server_app(base_dir=base_dir, folder_raw=folder_raw, maintenance=False)


app = MaintenanceGateway(factory=build_app, base=base_dir)
