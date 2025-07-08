from flask import Flask, render_template, request, flash, redirect, url_for
import csv
import os
from datetime import datetime
import threading
import time

app = Flask(__name__)
app.secret_key = 'your-secret-key-here'

CSV_FILE_PATH = r'\\lofs010\cmip_groups\LOG-DGQ\3_PROCESSOS E PRODUTO\02 - Controlo de Produto\00 - Equipa\07 - Catarina_Pinheiro\form_data.csv' 

# Lock for thread-safe file operations
file_lock = threading.Lock()

def save_to_csv(data):
    """Save form data to CSV file - Windows compatible"""
    max_retries = 3
    retry_delay = 1  # seconds
    
    for attempt in range(max_retries):
        try:
            with file_lock:
                # Add timestamp
                data['Timestamp'] = datetime.now().strftime('%Y-%m-%d %H:%M:%S')
                data['Submitted_From'] = request.environ.get('REMOTE_ADDR', 'Unknown')
                
                # Check if file exists to determine if we need headers
                file_exists = os.path.exists(CSV_FILE_PATH)
                
                # Ensure directory exists
                os.makedirs(os.path.dirname(CSV_FILE_PATH), exist_ok=True)
                
                # Open file in append mode
                with open(CSV_FILE_PATH, 'a', newline='', encoding='utf-8') as csvfile:
                    fieldnames = ['Full Name', 'Email', 'Age', 'Gender', 'Message', 'Timestamp', 'Submitted_From']
                    writer = csv.DictWriter(csvfile, fieldnames=fieldnames)
                    
                    # Write header if file doesn't exist
                    if not file_exists:
                        writer.writeheader()
                    
                    # Write data
                    writer.writerow(data)
                    
                    # Force write to disk
                    csvfile.flush()
                    os.fsync(csvfile.fileno())
                
                print(f"Successfully saved data to {CSV_FILE_PATH}")
                return True
                
        except Exception as e:
            print(f"Attempt {attempt + 1} failed: {e}")
            if attempt < max_retries - 1:
                time.sleep(retry_delay)
            else:
                print(f"All attempts failed. Final error: {e}")
                return False
    
    return False

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # Get form data
        form_data = {
            'Full Name': request.form.get('fullname', ''),
            'Email': request.form.get('email', ''),
            'Age': request.form.get('age', ''),
            'Gender': request.form.get('gender', ''),
            'Message': request.form.get('message', '')
        }
        
        print(f"Received form data: {form_data}")
        print(f"Attempting to save to: {CSV_FILE_PATH}")
        
        # Save to CSV
        if save_to_csv(form_data):
            flash('Form submitted and saved successfully!', 'success')
            print("Form saved successfully!")
        else:
            flash('Error saving form data. Please try again.', 'error')
            print("Failed to save form data!")
        
        return redirect(url_for('index'))
    
    return render_template('form.html')

if __name__ == '__main__':
    # Test file path on startup
    print(f"Starting application...")
    print(f"CSV file path: {CSV_FILE_PATH}")
    
    # Test if we can write to the directory
    try:
        test_dir = os.path.dirname(CSV_FILE_PATH)
        if not os.path.exists(test_dir):
            os.makedirs(test_dir, exist_ok=True)
            print(f"Created directory: {test_dir}")
        else:
            print(f"Directory exists: {test_dir}")
            
        # Test write permissions
        test_file = os.path.join(test_dir, 'test_write.txt')
        with open(test_file, 'w') as f:
            f.write('test')
        os.remove(test_file)
        print("Write permissions OK")
        
    except Exception as e:
        print(f"WARNING: Cannot write to {CSV_FILE_PATH}: {e}")
        print("Please check the path and permissions")
    
    # Run on all network interfaces
    app.run(host='0.0.0.0', port=5000, debug=True)