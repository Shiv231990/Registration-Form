from flask import Flask, render_template, request
import pandas as pd
import os

app = Flask(__name__)
EXCEL_FILE = 'data.xlsx'

@app.route('/')
def index():
    return render_template('form.html')

@app.route('/submit', methods=['POST'])
def submit():
    # Capture and clean the mobile number to prevent duplicate bypass
    mobile_input = str(request.form.get('mobile')).strip()
    
    # Organize form data into a dictionary
    user_data = {
        'Name': [request.form.get('username')],
        'Email': [request.form.get('email')],
        'Mobile': [mobile_input],
        'Address': [request.form.get('address')],
        'City': [request.form.get('city')],
        'Salary': [request.form.get('salary')]
    }
    
    df_new = pd.DataFrame(user_data)

    # Check if the Excel file already exists
    if os.path.exists(EXCEL_FILE):
        # engine='openpyxl' ensures compatibility with .xlsx files
        df_existing = pd.read_excel(EXCEL_FILE, engine='openpyxl')
        
        # Check if 'Mobile' column exists and look for duplicates
        if 'Mobile' in df_existing.columns:
            # Clean existing numbers for an accurate comparison
            existing_mobiles = df_existing['Mobile'].astype(str).str.strip().tolist()
            
            if mobile_input in existing_mobiles:
                # Return Bold Error Message
                return """
                <div style='text-align:center; margin-top:50px; font-family:sans-serif;'>
                    <h2 style='font-weight:bold; color:red;'>DETAILS ALREADY EXISTS!</h2>
                    <p style='font-weight:bold;'>THIS MOBILE NUMBER IS ALREADY REGISTERED.</p>
                    <a href='/' style='font-weight:bold; color:blue; text-decoration:none;'>GO BACK</a>
                </div>
                """
        
        # Append new data to existing data
        df_final = pd.concat([df_existing, df_new], ignore_index=True)
    else:
        # Create fresh data if file doesn't exist
        df_final = df_new

    # Save to Excel
    try:
        df_final.to_excel(EXCEL_FILE, index=False, engine='openpyxl')
        return """
        <div style='text-align:center; margin-top:50px; font-family:sans-serif;'>
            <h2 style='font-weight:bold; color:green;'>SUCCESS!</h2>
            <p style='font-weight:bold;'>DATA SAVED TO EXCEL SUCCESSFULLY.</p>
            <a href='/' style='font-weight:bold; color:blue; text-decoration:none;'>SUBMIT ANOTHER</a>
        </div>
        """
    except PermissionError:
        return "<h2 style='font-weight:bold; color:red;'>ERROR: PLEASE CLOSE THE EXCEL FILE AND TRY AGAIN!</h2>"

if __name__ == '__main__':
    app.run(debug=True)
