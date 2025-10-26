from flask import Flask, jsonify
from flask_cors import CORS
import pandas as pd
import os

app = Flask(__name__)
CORS(app)

EXCEL_FILE = 'data.xlsx'

@app.route('/')
def home():
    return jsonify({
        'message': 'Excel Data API is running',
        'status': 'ok'
    })

@app.route('/data')
def get_data():
    try:
        # Read Excel and convert EVERYTHING to string
        df = pd.read_excel(EXCEL_FILE, dtype=str, na_filter=False)
        
        # Convert to JSON
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

@app.route('/sheets')
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
