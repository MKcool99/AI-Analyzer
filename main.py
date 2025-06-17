from flask import Flask, render_template, request, flash, redirect, url_for
import os
import pandas as pd
from werkzeug.utils import secure_filename
from openai import OpenAI
from dotenv import load_dotenv

load_dotenv()

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 5 * 1024 * 1024 * 1024  
app.config['SECRET_KEY'] = 'your-secret-key-here'  

client = OpenAI()


ALLOWED_EXTENSIONS = {'xls', 'xlsx'}


def allowed_file(filename):
    """Check if the uploaded file has an allowed extension"""
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


def create_data_summary(df, file_info):
    """Create a concise summary of the Excel data for the AI prompt"""
    summary = f"""
Data Summary:
- File: {file_info['filename']}
- Total Rows: {file_info['rows']}
- Total Columns: {file_info['columns']}
- Column Names: {', '.join(file_info['column_names'])}
"""
    
    if file_info.get('numeric_stats'):
        summary += "\nNumeric Column Statistics:\n"
        for col, stats in file_info['numeric_stats'].items():
            summary += f"- {col}: Sum={stats['sum']:.2f}, Average={stats['mean']:.2f}, Count={stats['count']}\n"
    
    if len(df) <= 20:
        summary += f"\nFirst few rows of data:\n{df.head(10).to_string()}"
    else:
        summary += f"\nSample data (first 5 rows):\n{df.head(5).to_string()}"
    
    return summary


def ask_openai_about_data(data_summary, user_question, model="gpt-3.5-turbo", temperature=0.7, max_tokens=500):
    """Ask OpenAI to analyze the data and answer the user's question"""
    try:
        prompt = f"""You are a data analyst assistant. I will provide you with information about an Excel spreadsheet and a user's question about that data.

{data_summary}

User's Question: {user_question}

Please analyze the data provided above and answer the user's question. Be specific and reference the actual data values when possible. If the question cannot be fully answered with the provided data, explain what information is available and what might be missing."""

        chat_completion = client.chat.completions.create(
            model=model,
            messages=[
                {"role": "system", "content": "You are a helpful data analyst assistant specializing in Excel data analysis."},
                {"role": "user", "content": prompt}
            ],
            temperature=temperature,
            max_tokens=max_tokens,
        )
        return chat_completion.choices[0].message.content.strip()
    except Exception as e:
        return f"Error analyzing data: {str(e)}"


@app.route('/')
def index():
    """Render the main page"""
    return render_template('index.html')


@app.route('/upload', methods=['GET', 'POST'])
def upload_file():
    """Handle file upload and question processing"""
    if request.method == 'GET':
        return redirect(url_for('index'))
    
    if 'file' not in request.files:
        flash('No file selected')
        return redirect(url_for('index'))
    
    file = request.files['file']
    query = request.form.get('question', '').strip()
    
    if file.filename == '':
        flash('No file selected')
        return redirect(url_for('index'))
    
    if not query:
        flash('Please provide a question about the file')
        return redirect(url_for('index'))
    
    if not allowed_file(file.filename):
        flash('Invalid file type. Please upload .xls or .xlsx files only.')
        return redirect(url_for('index'))
    
    try:
        filename = secure_filename(file.filename)
        
        print(f"File received: {filename}")
        print(f"Question received: {query}")
        
        file_extension = filename.rsplit('.', 1)[1].lower()
        engine = 'openpyxl' if file_extension == 'xlsx' else 'xlrd'
        
        df = pd.read_excel(file, engine=engine)
        
        numeric_stats = {}
        for col in df.select_dtypes(include=['number']).columns:
            numeric_stats[col] = {
                'sum': df[col].sum(),
                'mean': df[col].mean(),
                'count': df[col].count()
            }
        
        try:
            xl_file = pd.ExcelFile(file, engine=engine)
            sheet_names = xl_file.sheet_names
        except:
            sheet_names = ['Sheet1'] 
        
        file_info = {
            'filename': filename,
            'rows': len(df),
            'columns': len(df.columns),
            'column_names': list(df.columns),
            'sheet_names': sheet_names,
            'numeric_stats': numeric_stats
        }
        
        data_summary = create_data_summary(df, file_info)
        print("Sending request to OpenAI...")
        ai_answer = ask_openai_about_data(data_summary, query)
        print(f"AI Response: {ai_answer}")
        
        display_df = df.head(100) if len(df) > 100 else df
        file_data = display_df.to_html(classes='table table-bordered table-striped', 
                                      index=False, 
                                      table_id='data-table')
        
        return render_template('index.html', 
                             file_data=file_data, 
                             query=query, 
                             file_info=file_info,
                             ai_answer=ai_answer,
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
