#!/bin/bash
python3 -m venv pc2Automation
source ./pc2Automation/bin/activate
./pc2Automation/bin/pip3 install --upgrade pip
./pc2Automation/bin/pip3 install Cython
./pc2Automation/bin/pip3 install pandas
./pc2Automation/bin/pip3 install ib_insync
./pc2Automation/bin/pip3 install openpyxl
