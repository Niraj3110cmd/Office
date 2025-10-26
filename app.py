from flask import Flask, jsonify, request, Response
from flask_cors import CORS
import pandas as pd
import os
import json
import numpy as np

app = Flask(__name__)
CORS(app)

EXCEL_FILE = 'data.xlsx'

def safe_convert(val):
    """Safely convert values to JSON-compatible format"""
    # Handle None/NaN/NaT
    if val is None or pd.isna(val):
        return None
    
    # Handle integers
    if isinstance(val, (np.integer, int)):
        return int(val)
    
    # Handle floats - this is the key fix
    if isinstance(val, (np.floating, float)):
        # Check for NaN or infinity
        if np.isnan(val) or np.isinf(val):
            return None
        # Check if value is too large for JSON
        try:
            # Try to convert to float
            float_val = float(val)
            # Check if it's a reasonable number
            if abs(float_val) > 1e308:  # Max float in JSON
                return None
            return round(float_val, 2)
        except (ValueError, OverflowError):
            return None
    
    # Handle booleans
    if isinstance(val, (bool, np.bool_)):
        return bool(val)
    
    # Handle timestamps
    if isinstance(val, pd.Timestamp):
        return val.strftime('%Y-%m-%d %H:%M:%S')
    
    # Convert everything else to string
    try:
        str_val = str(val).strip()
        if str_val.lower() in ['nan', 'none', 'nat', '']:
            return None
        return str_val
    except:
        return None

@app.route('/')
def home():
    return jsonify({
        'message': 'Excel Data API is running',
        'status': 'ok',
        'endpoints': {
            '/data': 'Get all data as JSON',
            '/sheets': 'Get list of sheet names'
        }
    })

@app.route('/data', methods=['GET'])
def get_data():
    try:
        sheet_name = request.args.get('sheet', None)
        
        # Read Excel
        if sheet_name:
            df = pd.read_excel(EXCEL_FILE, sheet_name=sheet_name)
        else:
            df = pd.read_excel(EXCEL_FILE)
        
        # Clean column names
        df.columns = [str(col).strip() for col in df.columns]
        
        # Apply safe conversion to all cells
        for col in df.columns:
            df[col] = df[col].apply(safe_convert)
        
        # Convert to list of dictionaries
        records = df.to_dict(orient='records')
        
        response_data = {
            'success': True,
            'rows': len(records),
            'columns': list(df.columns),
            'data': records
        }
        
        # Use json.dumps with specific settings
        json_str = json.dumps(
            response_data,
            ensure_ascii=False,
            allow_nan=False,
            indent=None
        )
        
        return Response(
            json_str,
            mimetype='application/json',
            headers={'Content-Type': 'application/json; charset=utf-8'}
        )
    
    except Exception as e:
        return jsonify({
            'success': False,
            'error': str(e),
            'type': type(e).__name__
        }), 400

@app.route('/sheets', methods=['GET'])
def get_sheets():
    try:
        xl_file = pd.ExcelFile(EXCEL_FILE)
        return jsonify({
            'success': True,
            'sheets': xl_file.sheet_names
        })
    except Exception as e:
        return jsonify({
            'success': False,
            'error': str(e)
        }), 400

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
