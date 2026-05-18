lines = open('monthly_epic_plan_progress_service.py', encoding='utf-8').readlines()
# Find _approved_dates function
for i, l in enumerate(lines):
    if 'def _approved_dates' in l:
        print(f'=== _approved_dates at line {i+1} ===')
        for j in range(i, min(i+18, len(lines))):
            print(f'{j+1}: {lines[j]}', end='')
        print()
    if 'epic_plan = ((planner_row' in l:
        print(f'=== epic_plan assignment at line {i+1} ===')
        for j in range(max(0,i-5), min(i+30, len(lines))):
            print(f'{j+1}: {lines[j]}', end='')
        print()
    if '_approved_dates(epic_plan)' in l:
        print(f'=== _approved_dates call at line {i+1} ===')
        for j in range(max(0,i-3), min(i+5, len(lines))):
            print(f'{j+1}: {lines[j]}', end='')
