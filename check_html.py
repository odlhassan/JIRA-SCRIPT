with open('report_html/team_capacity_planner.html', encoding='utf-8') as f:
    content = f.read()
print(f'File size: {len(content)} chars')
print('Modal HTML present:', 'id="jira-sync-modal"' in content)
print('Checkboxes present:', 'jira-row-chk' in content)
print('Toolbar present:', 'jira-sel-toolbar' in content)
print('Select Unsynced btn:', 'sel-unsynced-btn' in content)
print('_jiraSyncSelectedIds:', '_jiraSyncSelectedIds' in content)
print('ids sent in push:', 'JSON.stringify({ ids })' in content)
