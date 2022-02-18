#!/bin/bash
python3 -m venv pc2Automation
source ./pc2Automation/bin/activate
./pc2Automation/bin/pip install --no-cache-dir "numpy==1.22.0"
./pc2Automation/bin/pip install --no-cache-dir "pandas==1.3.5"
./pc2Automation/bin/pip install --no-cache-dir "openpyxl==3.0.9"