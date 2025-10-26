from flask import Flask, Response
from flask_cors import CORS
import pandas as pd
import json
import os
import numpy as np

app = Flask(__name__)
CORS(app)

EXCEL_FILE = 'data.xlsx'

@app.route('/')
def home():
    return Response(
        json.dumps({'message': 'API running', 'status': 'ok'}),
        mimetype='application/json'
    )

@app.route('/data')
def get_data():
    try:
        df = pd.read_excel(EXCEL_FILE, sheet_name=0)
        
        clean_data = []
        
        for index, row in df.iterrows():
            clean_row = {}
            for col in df.columns:
                value = row[col]
                
                if pd.isna(value) or value == 'NA' or str(value).strip() == 'NA':
                    clean_row[str(col)] = None
                elif isinstance(value, (int, np.integer)):
                    clean_row[str(col)] = int(value)
                elif isinstance(value, (float, np.floating)):
                    if np.isnan(value) or np.isinf(value):
                        clean_row[str(col)] = None
                    else:
                        clean_row[str(col)] = round(float(value), 2)
                elif isinstance(value, pd.Timestamp):
                    clean_row[str(col)] = value.strftime('%Y-%m-%d %H:%M:%S')
                else:
                    str_value = str(value).strip()
                    if str_value in ['', 'nan', 'None', 'NaT', 'NA']:
                        clean_row[str(col)] = None
                    else:
                        clean_row[str(col)] = str_value
            
            clean_data.append(clean_row)
        
        response = {
            'success': True,
            'rows': len(clean_data),
            'data': clean_data
        }
        
        return Response(
            json.dumps(response, ensure_ascii=False),
            mimetype='application/json'
        )
        
    except Exception as e:
        return Response(
            json.dumps({'success': False, 'error': str(e)}),
            status=500,
            mimetype='application/json'
        )

@app.route('/sheets')
def get_sheets():
    try:
        xl_file = pd.ExcelFile(EXCEL_FILE)
        response = {'success': True, 'sheets': xl_file.sheet_names}
        return Response(
            json.dumps(response),
            mimetype='application/json'
        )
    except Exception as e:
        return Response(
            json.dumps({'success': False, 'error': str(e)}),
            status=500,
            mimetype='application/json'
        )

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
