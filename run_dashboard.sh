#!/bin/bash
cd "$(dirname "$0")"
pip3 install -q pandas numpy openpyxl streamlit
streamlit run dashboard.py
