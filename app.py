from flask import Flask, jsonify, request
from flask_cors import CORS
import pandas as pd
import os
import json
import numpy as np

app = Flask(__name__)
CORS(app)

EXCEL_FILE = 'data.xlsx'

def clean_value(val):
    """Clean individual values for JSON serialization"""
    if pd.isna(val):
        return None
    if isinstance(val, (np.integer, int)):
        return int(val)
    if isinstance(val, (np.floating, float)):
        if np.isinf(val) or np.isnan(val):
            return None
        return float(val)
    # Convert to string and remove any non-printable characters
    val_str = str(val)
    # Remove any characters that might cause JSON issues
    cleaned = ''.join(char for char in val_str if char.isprintable() or char in '\n\r\t')
    return cleaned.strip()

def clean_dataframe(df):
    """Clean DataFrame for JSON serialization"""
    # Create a new dataframe with cleaned values
    cleaned_df = pd.DataFrame()
    
    for col in df.columns:
        # Clean column name
        clean_col_name = ''.join(char for char in str(col) if char.isprintable()).strip()
        
        # Clean all values in the column
        cleaned_df[clean_col_name] = df[col].apply(clean_value)
    
    return cleaned_df

@app.route('/')
def home():
    return jsonify({
        'message': 'Excel Data API is running',
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
        
        # Clean the DataFrame
        df = clean_dataframe(df)
        
        # Convert to records
        data = df.to_dict(orient='records')
        
        # Return with explicit JSON serialization
        response = {
            'success': True,
            'rows': len(data),
            'data': data
        }
        
        return app.response_class(
            response=json.dumps(response, ensure_ascii=False),
            status=200,
            mimetype='application/json'
        )
    
    except Exception as e:
        return jsonify({
            'success': False,
            'error': str(e)
        }), 400

@app.route('/sheets', methods=['GET'])
def get_sheets():
    try:
        xl_file = pd.ExcelFile(EXCEL_FILE)
        sheet_names = xl_file.sheet_names
        
        return jsonify({
            'success': True,
            'sheets': sheet_names
        })
    
    except Exception as e:
        return jsonify({
            'success': False,
            'error': str(e)
        }), 400

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
