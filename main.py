from flask import Flask, render_template, request, flash, redirect, url_for
import os
import pandas as pd
from werkzeug.utils import secure_filename

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 5 * 1024 * 1024 * 1024  # we dont want to overload our systems so i added a 5gb limit
app.config['SECRET_KEY'] = 'your-secret-key-here'  #need this fo rflash msessages

#to prevent any issues and hacking (though unreasonable) we limit the amount of file extentions
ALLOWED_EXTENSIONS = {'xls', 'xlsx'}

def allowed_file(filename):
    """Check if the uploaded file has an allowed extension"""
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/')
def index():
    """Render the main page"""
    return render_template('index.html')

@app.route('/upload', methods=['GET', 'POST'])
def upload_file():
    """Handle file upload and question processing"""
    if request.method == 'GET':
        return redirect(url_for('index'))
    
    #check if file and question were given
    if 'file' not in request.files:
        flash('No file selected')
        return redirect(url_for('index'))
    
    file = request.files['file']
    query = request.form.get('question', '').strip()
    
    #check file selection
    if file.filename == '':
        flash('No file selected')
        return redirect(url_for('index'))
    
    #validate question
    if not query:
        flash('Please provide a question about the file')
        return redirect(url_for('index'))
    
    #check the file type (again for cybersecurity)
    if not allowed_file(file.filename):
        flash('Invalid file type. Please upload .xls or .xlsx files only.')
        return redirect(url_for('index'))
    
    try:
        #secure filename
        filename = secure_filename(file.filename)
        
        #log down recieved data
        print(f"File received: {filename}")
        print(f"Question received: {query}")
        
        #look at the correct engine for pandas
        file_extension = filename.rsplit('.', 1)[1].lower()
        engine = 'openpyxl' if file_extension == 'xlsx' else 'xlrd'
        
        #read the file
        df = pd.read_excel(file, engine=engine)
        
        #df to html for display
        file_data = df.to_html(classes='table table-bordered table-striped', 
                              index=False, 
                              table_id='data-table')
        
        #get file info
        file_info = {
            'filename': filename,
            'rows': len(df),
            'columns': len(df.columns),
            'column_names': list(df.columns)
        }
        
        return render_template('index.html', 
                             file_data=file_data, 
                             query=query, 
                             file_info=file_info,
                             success=True)
    
    except Exception as e:
        print(f"Error processing file: {str(e)}")
        flash(f'Error processing file: {str(e)}')
        return redirect(url_for('index'))

@app.errorhandler(413)
def too_large(e):
    """Handle file too large error"""
    flash("File is too large. Maximum size is 5GB.")
    return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=5000, debug=True)
