from flask import Flask, render_template, send_file, request, jsonify, session, redirect, url_for
import pandas as pd
from datetime import datetime
from flask_cors import CORS

app = Flask(__name__)
CORS(app)

app.secret_key = "your_secret_key"  # Set a secret key for session management

# User credentials (in a real-world application, these should be stored securely)
valid_credentials = {'username': 'numericx', 'password': 'numericx@1234!'}

# Load the Excel files into pandas DataFrames
quote_file = 'quote.xlsx'
pop_file = 'popup.xlsx'

try:
    quoteFile = pd.read_excel(quote_file)
except FileNotFoundError:
    quoteFile = pd.DataFrame(columns=[
        'Index', 'Timestamp', 'Name', 'BusinessName', 'Email', 'BusinessType',
        'Phone', 'AnnualTurnover', 'CompanyRegistration', 'Bookkeeping',
        'VATReturns', 'Payroll', 'PayslipsPerMonth', 'ServiceDuration', 'QuotePrice'
    ]) 

try:
    popupFile = pd.read_excel(pop_file)
except FileNotFoundError:
    popupFile = pd.DataFrame(columns=[
        'Index', 'Timestamp', 'Name', 'Email', 'Phone', 'ServiceType', 'Message'
    ])

@app.route('/dashboard')
def index():
    return render_template('dashboard.html')

@app.route('/popup-records')
def popup_screen():
    # Reverse the order of records
    reversed_records = popupFile.iloc[::-1].to_dict('records')
    # Pass the reversed records to the template for rendering
    return render_template('popup.html', records=reversed_records)

@app.route('/quote-records')
def quote_screen():
    # Reverse the order of records
    reversed_records = quoteFile.iloc[::-1].to_dict('records')
    # Pass the reversed records to the template for rendering
    return render_template('quote.html', records=reversed_records)

@app.route('/download_quote')
def download_quote():
    # Serve the Excel file for download
    return send_file(quote_file, as_attachment=True)

@app.route('/download_popup')
def download_popup():
    # Serve the Excel file for download
    return send_file(pop_file, as_attachment=True)

@app.route('/add_popup_record', methods=['POST'])
def add_popup_record():
    global popupFile  # Declare popupFile as global to modify the global variable

    # Extract JSON data from the request
    data = request.json

    # Prepare data for adding to DataFrame
    new_record = {
        'Index': len(popupFile) + 1,
        'Timestamp': datetime.now(),
        'Name': data.get('name', ''),
        'Email': data.get('email', ''),
        'Phone': data.get('phone', ''),
        'ServiceType': data.get('serviceType', ''),
        'Message': data.get('message', '')
    }
    # Convert the data into a DataFrame
    new_record_df = pd.DataFrame([new_record])

    # Concatenate the new record with the existing DataFrame
    popupFile = pd.concat([popupFile, new_record_df], ignore_index=True)

    # Write the updated DataFrame back to the Excel file
    popupFile.to_excel(pop_file, index=False)

    # Return success response
    return jsonify({'message': 'Record added successfully'}), 201

@app.route('/add_quote_record', methods=['POST'])
def add_quote_record():
    global quoteFile  # Declare quoteFile as global to modify the global variable

    # Extract JSON data from the request
    data = request.json

    # Prepare data for adding to DataFrame
    new_record = {
        'Index': len(quoteFile) + 1,
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
    quoteFile = pd.concat([quoteFile, new_record_df], ignore_index=True)

    # Write the updated DataFrame back to the Excel file
    quoteFile.to_excel(quote_file, index=False)

    # Return success response
    return jsonify({'message': 'Record added successfully'}), 201

@app.route('/delete_quote_record', methods=['POST'])
def delete_quote_record():
    global quoteFile  # Declare quoteFile as global to modify the global variable

    # Extract index of the record to delete from the request
    index = int(request.json['index'])

    # Delete the record from the DataFrame
    quoteFile.drop(index, inplace=True)

    # Write the updated DataFrame back to the Excel file
    quoteFile.to_excel(quote_file, index=False)

    # Return success response
    return jsonify({'message': 'Record deleted successfully'}), 200

@app.route('/delete_popup_record', methods=['POST'])
def delete_popup_record():
    global popupFile  # Declare popupFile as global to modify the global variable

    # Extract index of the record to delete from the request
    index = int(request.json['index'])

    # Delete the record from the DataFrame
    popupFile.drop(index, inplace=True)

    # Write the updated DataFrame back to the Excel file
    popupFile.to_excel(pop_file, index=False)

    # Return success response
    return jsonify({'message': 'Record deleted successfully'}), 200

@app.route('/')
@app.route('/login', methods=['GET', 'POST'])
def login():
    if request.method == 'POST':
        username = request.form['username']
        password = request.form['password']
        if username == valid_credentials['username'] and password == valid_credentials['password']:
            session['username'] = username
            return redirect(url_for('index'))  # Update this line
        else:
            return render_template('login.html', message='Invalid credentials. Please try again.')
    return render_template('login.html')

@app.route('/logout')
def logout():
    session.pop('username', None)
    return redirect(url_for('login'))

if __name__ == '__main__':
    app.run(debug=True)
