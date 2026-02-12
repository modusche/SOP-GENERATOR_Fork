"""
Headless SOP Generator server - no GUI window.
Used by the Camunda Modeler plugin to generate Word documents.
"""
import sys
import os

# Ensure we can find modules relative to this script
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from app import app
from waitress import serve

if __name__ == '__main__':
    port = int(sys.argv[1]) if len(sys.argv) > 1 else 8000
    print(f'SOP Generator server starting on http://127.0.0.1:{port}')
    serve(app, host='127.0.0.1', port=port, _quiet=True)
