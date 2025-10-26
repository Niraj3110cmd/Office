from flask import Flask, Response
from flask_cors import CORS
import pandas as pd
import json
import os

app = Flask(__name__)
CORS(app)

EXCEL_FILE = 'data.xlsx'

@app.route('/')
def home():
    return Response(
        json.dumps({'message': 'API is running', 'status': 'ok'}),
        mimetype='application/json'
    )

@app.route('/data')
def get_data():
    try:
        # Read Excel - convert everything to string immediately
        df = pd.read_excel(EXCEL_FILE)
        
        # Force convert ALL columns to string, replacing any problematic values
        for col in df.columns:
            df[col] = df[col].astype(str).replace(['nan', 'None', 'NaT', 'inf', '-inf'], '')
        
        # Clean column names too
        df.columns = [str(col).strip() for col in df.columns]
        
        # Convert to records
        records = df.to_dict(orient='records')
        
        # Build response manually
        response = {
            'success': True,
            'rows': len(records),
            'data': records
        }
        
        # Manually serialize to JSON with strict settings
        json_output = json.dumps(response, ensure_ascii=True, indent=None)
        
        return Response(
            json_output,
            mimetype='application/json',
            headers={'Content-Type': 'application/json'}
        )
        
    except Exception as e:
        error = {'success': False, 'error': str(e)}
        return Response(
            json.dumps(error),
            status=400,
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
        error = {'success': False, 'error': str(e)}
        return Response(
            json.dumps(error),
            status=400,
            mimetype='application/json'
        )

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
