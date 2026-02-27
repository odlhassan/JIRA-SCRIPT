from __future__ import annotations

import argparse
import os
import socket
from pathlib import Path

from report_server import run_report_server, sync_report_html


def _clear_cache_files(base_dir: Path) -> None:
    cache_files = ("jira_sync_cache.db", "run_all_state.json")
    for file_name in cache_files:
        file_path = base_dir / file_name
        try:
            file_path.unlink(missing_ok=True)
            print(f"[server] Cleared cache file: {file_path.name}")
        except OSError as exc:
            print(f"[server] Warning: failed to remove {file_path.name}: {exc}")


def _port_is_available(host: str, port: int) -> bool:
    try:
        with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
            sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            sock.bind((host, port))
        return True
    except OSError:
        return False


def _resolve_server_port(host: str, requested_port: int) -> int:
    if _port_is_available(host, requested_port):
        return requested_port

    if requested_port == 3000:
        for candidate in (3001, 3002, 3003):
            if _port_is_available(host, candidate):
                print(f"[server] Port 3000 is busy; switching to {candidate}.")
                return candidate
        raise RuntimeError("[server] Ports 3000-3003 are all in use. Stop one and retry.")

    raise RuntimeError(f"[server] Port {requested_port} is busy. Choose a different --port.")


def main() -> None:
    parser = argparse.ArgumentParser(
        description="Start report server without running data generation scripts."
    )
    parser.add_argument(
        "--report-html-dir",
        default=os.getenv("REPORT_HTML_DIR", "report_html"),
        help="Directory containing generated report HTML files.",
    )
    parser.add_argument(
        "--host",
        default=os.getenv("REPORT_HOST", "127.0.0.1"),
        help="Host for local report server.",
    )
    parser.add_argument(
        "--port",
        type=int,
        default=int(os.getenv("REPORT_PORT", os.getenv("PORT", "3000"))),
        help="Port for local report server.",
    )
    parser.add_argument(
        "--fresh",
        action="store_true",
        help="Delete local sync cache/state files before starting the server.",
    )
    parser.add_argument(
        "--no-sync",
        action="store_true",
        help="Do not sync generated HTML files into report_html before starting the server.",
    )
    args = parser.parse_args()

    base_dir = Path(__file__).resolve().parent
    if args.fresh:
        _clear_cache_files(base_dir)
    if args.no_sync:
        print("[server] Skipping report-html sync (--no-sync).")
    else:
        moved = sync_report_html(base_dir, args.report_html_dir)
        if moved == 0:
            print("[server] report-html sync completed: no files moved.")
        else:
            print(f"[server] report-html sync completed: {moved} file(s) moved.")

    port = _resolve_server_port(args.host, args.port)

    run_report_server(
        base_dir=base_dir,
        folder_raw=args.report_html_dir,
        host=args.host,
        port=port,
    )


if __name__ == "__main__":
    main()
