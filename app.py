from flask import Flask, render_template, send_file, request, jsonify
import pandas as pd
from datetime import datetime
from flask_cors import CORS

app = Flask(__name__)
CORS(app)  # Enable CORS for all origins

# Load the Excel file into a pandas DataFrame
excel_file = 'records.xlsx'
try:
    df = pd.read_excel(excel_file)
except FileNotFoundError:
    df = pd.DataFrame(columns=['Timestamp', 'Counter', 'Column1', 'Column2'])  # Create a new DataFrame if file doesn't exist

@app.route('/')
def index():
    # Reverse the order of records
    reversed_records = df.iloc[::-1].to_dict('records')
    # Pass the reversed records to the template for rendering
    return render_template('index.html', records=reversed_records)

@app.route('/download')
def download():
    # Serve the Excel file for download
    return send_file(excel_file, as_attachment=True)

@app.route('/add_record', methods=['POST'])
def add_record():
    # Declare df as a global variable
    global df

    # Extract JSON data from the request
    data = request.json

    # Prepare data for adding to DataFrame
    new_record = {
        'Index': len(df) + 1,
        'Timestamp': datetime.now(),
        'Name': data.get('name', ''),
        'BusinessName': data.get('businessName', ''),
        'Email': data.get('email', ''),
        'BusinessType': data.get('businessType', ''),
        'Phone': data.get('phone', ''),
        'AnnualTurnover': data.get('annualTurnover', ''),
        'CompanyRegistration': data.get('companyRegistration', ''),
        'Bookkeeping': data.get('bookkeeping', ''),
        'VATReturns': data.get('vatReturns', ''),
        'Payroll': data.get('payroll', ''),
        'PayslipsPerMonth': data.get('payslipsPerMonth', ''),
        'ServiceDuration': data.get('serviceDuration', ''),
        'QuotePrice': data.get('quotePrice', '')
    }

    # Convert the data into a DataFrame
    new_record_df = pd.DataFrame([new_record])

    # Concatenate the new record with the existing DataFrame
    df = pd.concat([df, new_record_df], ignore_index=True)

    # Write the updated DataFrame back to the Excel file
    df.to_excel(excel_file, index=False)

    # Return success response
    return jsonify({'message': 'Record added successfully'}), 201


@app.route('/delete_record', methods=['POST'])
def delete_record():
    # Declare df as a global variable
    global df

    # Extract index of the record to delete from the request
    index = int(request.json['index'])

    # Delete the record from the DataFrame
    df.drop(index, inplace=True)

    # Write the updated DataFrame back to the Excel file
    df.to_excel(excel_file, index=False)

    # Return success response
    return jsonify({'message': 'Record deleted successfully'}), 200

if __name__ == '__main__':
    app.run(debug=True)
