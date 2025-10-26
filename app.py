from flask import Flask, jsonify, request, Response
from flask_cors import CORS
import pandas as pd
import os
import json
import numpy as np
import math

app = Flask(__name__)
CORS(app)

EXCEL_FILE = 'data.xlsx'

def is_valid_value(val):
    """Check if value is valid for JSON"""
    if val is None:
        return True
    if isinstance(val, (int, str, bool)):
        return True
    if isinstance(val, float):
        return not (math.isnan(val) or math.isinf(val))
    return True

def clean_value(val):
    """Clean individual values for JSON serialization"""
    # Handle None/NaN
    if val is None or pd.isna(val):
        return None
    
    # Handle numeric types
    if isinstance(val, (np.integer, int)):
        return int(val)
    
    if isinstance(val, (np.floating, float)):
        if math.isnan(val) or math.isinf(val):
            return None
        return round(float(val), 2)
    
    # Handle boolean
    if isinstance(val, (bool, np.bool_)):
        return bool(val)
    
    # Convert everything else to string
    try:
        val_str = str(val)
        # Remove non-printable characters except newlines and tabs
        cleaned = ''.join(char for char in val_str if char.isprintable() or char in '\n\r\t')
        cleaned = cleaned.strip()
        
        # Don't return empty strings
        if cleaned == '' or cleaned.lower() == 'nan' or cleaned.lower() == 'none':
            return None
            
        return cleaned
    except:
        return None

@app.route('/')
def home():
    return jsonify({
        'message': 'Excel Data API is running',
        'status': 'ok',
        'endpoints': {
            '/data': 'Get all data as JSON',
            '/data?sheet=SheetName': 'Get specific sheet data',
            '/sheets': 'Get list of all sheet names'
        }
    })

@app.route('/data', methods=['GET'])
def get_data():
    try:
        sheet_name = request.args.get('sheet', None)
        
        if sheet_name:
            df = pd.read_excel(EXCEL_FILE, sheet_name=sheet_name)
        else:
            df = pd.read_excel(EXCEL_FILE)
        
        # Clean column names
        df.columns = [str(col).strip() for col in df.columns]
        
        # Clean all data
        for col in df.columns:
            df[col] = df[col].apply(clean_value)
        
        # Convert to dict
        records = df.to_dict(orient='records')
        
        # Build response
        response_data = {
            'success': True,
            'rows': len(records),
            'data': records
        }
        
        # Serialize to JSON string first to catch any issues
        json_string = json.dumps(response_data, ensure_ascii=False, allow_nan=False)
        
        return Response(
            json_string,
            mimetype='application/json',
            headers={'Content-Type': 'application/json; charset=utf-8'}
        )
    
    except Exception as e:
        error_response = {
            'success': False,
            'error': str(e)
        }
        return Response(
            json.dumps(error_response),
            status=400,
            mimetype='application/json'
        )

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
