import pandas as pd
import numpy as np
from datetime import datetime
import xlsxwriter

def create_blank_template(filename='oil_gas_analysis_template.xlsx'):
    workbook = xlsxwriter.Workbook(filename)
    
    # Format definitions
    header = workbook.add_format({
        'bold': True,
        'font_color': 'white',
        'bg_color': '#0D5BA6',
        'border': 1,
        'align': 'center'
    })
    
    subheader = workbook.add_format({
        'bold': True,
        'bg_color': '#C5D9F1',
        'border': 1,
        'align': 'center'
    })
    
    number_format = workbook.add_format({
        'num_format': '#,##0.00',
        'border': 1
    })
    
    percent_format = workbook.add_format({
        'num_format': '0.00%',
        'border': 1
    })
    
    border = workbook.add_format({
        'border': 1
    })
    
    # 1. Company Input Sheet
    ws_input = workbook.add_worksheet('Company Input')
    companies = ['Company A', 'Company B', 'Company C', 'Company D', 'Company E']
    headers = ['Ticker', 'Company Name', 'Market Cap', 'Enterprise Value', 'Production (BOE/d)', 
              'Oil %', 'Gas %', 'NGL %', 'Primary Regions']
    
    ws_input.write(0, 0, 'Company Basic Information', header)
    for col, header_text in enumerate(headers):
        ws_input.write(1, col, header_text, subheader)
        ws_input.set_column(col, col, 15)
    
    # 2. Operational Metrics
    ws_ops = workbook.add_worksheet('Operational')
    ops_headers = ['Company', 'Production (BOE/d)', 'YoY Growth %', 'Oil Mix %', 
                  'F&D Cost/BOE', 'Operating Cost/BOE', 'Reserve Life (Years)', 
                  'RRR %', '1P Reserves', '2P Reserves']
    
    ws_ops.write(0, 0, 'Operational Metrics', header)
    for col, header_text in enumerate(ops_headers):
        ws_ops.write(1, col, header_text, subheader)
        ws_ops.set_column(col, col, 15)
    
    # 3. Financial Metrics
    ws_fin = workbook.add_worksheet('Financial')
    fin_sections = {
        'Income Statement': ['Revenue', 'EBITDA', 'EBIT', 'Net Income', 'EPS'],
        'Balance Sheet': ['Total Assets', 'Total Debt', 'Net Debt', 'Equity', 'Working Capital'],
        'Cash Flow': ['Operating CF', 'Capex', 'Free CF', 'Dividends', 'Share Buybacks']
    }
    
    row = 0
    for section, metrics in fin_sections.items():
        ws_fin.write(row, 0, section, header)
        ws_fin.write(row + 1, 0, 'Company', subheader)
        for col, metric in enumerate(metrics, 1):
            ws_fin.write(row + 1, col, metric, subheader)
        row += 8
    
    # 4. Efficiency Metrics
    ws_eff = workbook.add_worksheet('Efficiency')
    eff_headers = ['Company', 'ROCE %', 'ROE %', 'ROIC %', 'Capital Efficiency', 
                  'Reinvestment Rate', 'FCF Yield %', 'Payout Ratio %']
    
    ws_eff.write(0, 0, 'Efficiency Metrics', header)
    for col, header_text in enumerate(eff_headers):
        ws_eff.write(1, col, header_text, subheader)
        ws_eff.set_column(col, col, 15)
    
    # 5. Valuation
    ws_val = workbook.add_worksheet('Valuation')
    val_headers = ['Company', 'EV/EBITDA', 'P/E', 'P/B', 'EV/2P Reserves', 
                  'EV/Daily Production', 'NAV/Share', 'Premium/Discount to NAV %']
    
    ws_val.write(0, 0, 'Valuation Metrics', header)
    for col, header_text in enumerate(val_headers):
        ws_val.write(1, col, header_text, subheader)
        ws_val.set_column(col, col, 15)
    
    # 6. Risk Assessment
    ws_risk = workbook.add_worksheet('Risk')
    risk_sections = {
        'Operational Risk': ['Geographic', 'Reserve Quality', 'Cost Structure'],
        'Financial Risk': ['Leverage', 'Interest Coverage', 'Liquidity'],
        'ESG Risk': ['Carbon Intensity', 'Water Usage', 'Safety Record']
    }
    
    row = 0
    for section, metrics in risk_sections.items():
        ws_risk.write(row, 0, section, header)
        ws_risk.write(row + 1, 0, 'Company', subheader)
        for col, metric in enumerate(metrics, 1):
            ws_risk.write(row + 1, col, metric, subheader)
        row += 8
    
    # 7. Peer Comparison
    ws_peer = workbook.add_worksheet('Peer Comparison')
    peer_metrics = ['Metric', 'Company A', 'Company B', 'Company C', 'Company D', 'Industry Avg']
    key_metrics = ['Production Growth %', 'Operating Margin %', 'ROCE %', 'FCF Yield %', 
                  'EV/EBITDA', 'Net Debt/EBITDA', 'Reserve Life', 'F&D Cost/BOE']
    
    ws_peer.write(0, 0, 'Peer Comparison Matrix', header)
    for col, header_text in enumerate(peer_metrics):
        ws_peer.write(1, col, header_text, subheader)
    for row, metric in enumerate(key_metrics, 2):
        ws_peer.write(row, 0, metric, border)
    
    # Add data validation where needed
    for ws in [ws_input, ws_ops, ws_fin, ws_eff, ws_val, ws_risk, ws_peer]:
        ws.protect('', {'insert_rows': True, 'format_cells': True, 'insert_columns': True})
    
    workbook.close()

if __name__ == "__main__":
    create_blank_template()
