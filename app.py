from flask import Flask, jsonify, request
from flask_cors import CORS
import pandas as pd
import os

app = Flask(__name__)
CORS(app)  # Enable CORS for Power BI access

# Path to your Excel file
EXCEL_FILE = 'data.xlsx'

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
        
        # Convert datetime/date/time columns to strings
        for col in df.columns:
            if df[col].dtype == 'datetime64[ns]':
                df[col] = df[col].dt.strftime('%Y-%m-%d %H:%M:%S')
            elif df[col].dtype == 'object':
                # Handle time objects
                df[col] = df[col].astype(str)
        
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
