#!/usr/bin/env python
import ast
import sys

try:
    with open('report_server.py', encoding='utf-8') as f:
        content = f.read()
    ast.parse(content)
    print('Syntax OK')
    sys.exit(0)
except SyntaxError as e:
    print(f"Syntax Error: {e}")
    sys.exit(1)
