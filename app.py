from flask import Flask, jsonify, request
from flask_cors import CORS
import pandas as pd
import os
import json
from decimal import Decimal
import numpy as np

app = Flask(__name__)
CORS(app)  # Enable CORS for Power BI access

# Path to your Excel file
EXCEL_FILE = 'data.xlsx'

def clean_dataframe(df):
    """Clean DataFrame for JSON serialization"""
    # Replace NaN and infinity values
    df = df.replace([np.inf, -np.inf], None)
    df = df.where(pd.notnull(df), None)
    
    # Convert all columns to appropriate types
    for col in df.columns:
        # Handle datetime columns
        if pd.api.types.is_datetime64_any_dtype(df[col]):
            df[col] = df[col].dt.strftime('%Y-%m-%d %H:%M:%S')
        # Handle timedelta columns
        elif pd.api.types.is_timedelta64_dtype(df[col]):
            df[col] = df[col].astype(str)
        # Handle numeric columns
        elif pd.api.types.is_numeric_dtype(df[col]):
            df[col] = df[col].astype(object).where(df[col].notnull(), None)
        # Handle object columns (including time objects)
        else:
            df[col] = df[col].apply(lambda x: str(x) if pd.notnull(x) and not isinstance(x, str) else x)
    
    return df

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
        # Get sheet name from query parameter (optional)
        sheet_name = request.args.get('sheet', None)
        
        if sheet_name:
            # Read specific sheet
            df = pd.read_excel(EXCEL_FILE, sheet_name=sheet_name)
        else:
            # Read first sheet by default
            df = pd.read_excel(EXCEL_FILE)
        
        # Clean the DataFrame
        df = clean_dataframe(df)
        
        # Convert DataFrame to JSON
        data = df.to_dict(orient='records')
        
        return jsonify({
            'success': True,
            'rows': len(data),
            'data': data
        })
    
    except Exception as e:
        return jsonify({
            'success': False,
            'error': str(e)
        }), 400

@app.route('/sheets', methods=['GET'])
def get_sheets():
    try:
        # Get all sheet names
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
