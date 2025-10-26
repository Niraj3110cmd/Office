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
    return {'message': 'API running', 'status': 'ok'}

@app.route('/data')
def get_data():
    try:
        # Read Excel - don't try to infer types
        df = pd.read_excel(EXCEL_FILE, sheet_name=0)
        
        # Process each row manually to handle all the NA values and mixed types
        clean_data = []
        
        for index, row in df.iterrows():
            clean_row = {}
            for col in df.columns:
                value = row[col]
                
                # Handle different value types
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
                    # Convert to string and clean
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
        
        # Return clean JSON
        return Response(
            json.dumps(response, ensure_ascii=False, indent=None),
            mimetype='application/json',
            headers={'Content-Type': 'application/json; charset=utf-8'}
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
        return {'success': True, 'sheets': xl_file.sheet_names}
    except Exception as e:
        return {'success': False, 'error': str(e)}

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port, debug=False)
```

4. **Commit changes**
5. **Wait 3 minutes** for deployment
6. Test: https://office-laa2.onrender.com/data

Then in Power BI:
```
let
    source = Json.Document(Web.Contents("https://office-laa2.onrender.com/data")),
    data = Table.FromRecords(source[data])
in
    data
