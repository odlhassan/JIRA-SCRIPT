from pathlib import Path
from generate_employee_capacity_utilization_report import _html as capacity_html

ROOT = Path(__file__).resolve().parent

def _html(payload=None):
    html = capacity_html(payload)
    html = html.replace('Employee Capacity &amp; Utilization', 'Team Worklog Hours')
    html = html.replace('Canonical Jira worklogs, capacity, leave, and assigned estimates for every employee.', 'Daily logged hours and RLT leave for each employee. Zero indicates no counted worklogs.')
    html = html.replace('initDetailDrawer();new MutationObserver(wireDetailCells).observe(table,{childList:true,subtree:true});wireDetailCells();', '')
    html = html.replace('new MutationObserver(applyColorFormatting).observe(table,{childList:true,subtree:true});', '')
    html = html.replace('exportButton.addEventListener(\'click\',downloadExcel);', "exportButton.addEventListener('click',downloadMatrix);")
    start = html.index('function render(){if(!month.value)return;')
    end = html.index('async function loadCanonical()', start)
    html = html[:start] + ROOT.joinpath('team_worklog_hours.js').read_text(encoding='utf-8-sig') + '\n' + html[end:]
    # Utilization coloring has no meaningful equivalent for a daily worklog matrix.
    html = html.replace('<div class="field"><span>Color formatting</span><details class="color-picker">', '<div class="field" hidden><span>Color formatting</span><details class="color-picker">')
    html = html.replace('</style></head>', '''.matrix{width:max-content;min-width:100%}.matrix th,.matrix td{min-width:86px;text-align:center;padding:10px 8px}.matrix th:first-child,.matrix td:first-child{position:sticky;left:0;z-index:3;min-width:190px;text-align:left;background:#fff}.matrix th:first-child{background:#153d6c;z-index:7}.matrix th small,.matrix-day small{display:block;font-size:10px;margin-top:4px}.matrix-day.has-hours{background:#eaf8ef}.matrix-day.no-hours{background:#fff2f2;color:#a12a2a}.matrix-day small{color:#88560a;font-weight:700}</style></head>''')
    return html

if __name__ == '__main__':
    ROOT.joinpath('team_worklog_hours_report.html').write_text(_html(), encoding='utf-8')
