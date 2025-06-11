from flask import Flask, render_template, request
import os
import pandas as pd

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH']=5000000000*1024 # Limit the upload size to 5GB
#Run the initial index.html file
@app.route('/')
def index():
    file_data = None
    return render_template('index.html', file_data=file_data)

# Configure the uploading of Excel files
@app.route('/upload', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        file = request.files['file']
        query = request.form.get('question')
        end = file.filename.split('.')[-1]
        engine = 'openpyxl' if end == 'xlsx' else 'xlrd'
        df = pd.read_excel(file, engine=engine)
        file_data = df.to_html(classes='table table-bordered', index=False)
    return render_template('index.html', file_data=file_data, query=query)

# Start the Flask application
if __name__ == '__main__':
  app.run(host='0.0.0.0', port=5000)

# To run the application, use the command: python main.py
