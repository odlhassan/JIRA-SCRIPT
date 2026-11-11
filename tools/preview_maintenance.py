"""Run only the intermission page without opening Jira or report databases."""
import argparse
from pathlib import Path
import sys

ROOT = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(ROOT))

from flask import Flask
from maintenance_service import MaintenanceGateway


def main():
    parser = argparse.ArgumentParser(description=__doc__)
    parser.add_argument('--port', type=int, default=3096)
    args = parser.parse_args()
    core = Flask(__name__)

    @core.get('/')
    def ready():
        return 'App available. This is the database-free maintenance preview. Open /maintenance to play.'

    core.wsgi_app = MaintenanceGateway(core.wsgi_app, base=ROOT)
    core.run(host='127.0.0.1', port=args.port, threaded=True)


if __name__ == '__main__':
    main()
