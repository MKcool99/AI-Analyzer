# to run the code first install requirements.txt
# then enter python main.py
# then go to the link the terminal gives you

from flask import Flask, render_template, request, flash, redirect, url_for, session, jsonify
import os
import pandas as pd
import numpy as np
from werkzeug.utils import secure_filename
from openai import OpenAI
from dotenv import load_dotenv
import openpyxl
from datetime import datetime
import re
import uuid
import pickle
import tempfile

"""Flask web application for querying Excel spreadsheets with GPT."""

load_dotenv()

app = Flask(__name__)
app.config['MAX_CONTENT_LENGTH'] = 5 * 1024 * 1024 * 1024  # 5 GB max upload
app.config['SECRET_KEY'] = 'your-secret-key-here'

file_store = {}
client = OpenAI(api_key=os.get_env('API_KEY'))
#im aware it's bad practice to hardcode the API key, but this is just a demo.

ALLOWED_EXTENSIONS = {'xls', 'xlsx'}

def allowed_file(filename):
    """Check if the uploaded file has an allowed extension."""
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def detect_data_errors(df):
    """Detect common data errors in the spreadsheet."""
    errors = []
    for col in df.columns:
        for idx, value in enumerate(df[col]):
            if pd.isna(value):
                continue
            if isinstance(value, str):
                excel_errors = ['#DIV/0!', '#N/A', '#NAME?', '#NULL!', '#NUM!', '#REF!', '#VALUE!']
                if any(error in str(value) for error in excel_errors):
                    errors.append(f"Excel error '{value}' found in column '{col}' at row {idx + 2}")
    numeric_cols = df.select_dtypes(include=[np.number]).columns
    for col in numeric_cols:
        missing_count = df[col].isna().sum()
        if missing_count > 0:
            errors.append(f"Column '{col}' has {missing_count} missing values")
    return errors

def ask_openai_with_enhanced_context(data_summary, user_question, errors, trends, qa_history=None, model="gpt-3.5-turbo"):
    """Send the user's question and data summary to OpenAI with extra context."""
    try:
        context_additions = []
        if errors:
            context_additions.append(f"IMPORTANT: This spreadsheet contains {len(errors)} data quality issues that may affect analysis accuracy.")
        if trends:
            trend_summary = "Key trends identified: " + ", ".join([f"{t['column']} is {t['direction']}" for t in trends[:3]])
            context_additions.append(trend_summary)
        question_lower = user_question.lower()
        if any(word in question_lower for word in ['trend', 'growth', 'increase', 'decrease', 'change']):
            context_additions.append("The user is asking about trends - focus on the trend analysis data provided above.")
        if any(word in question_lower for word in ['error', 'problem', 'issue', 'wrong']):
            context_additions.append("The user is asking about data quality - reference the data quality issues section above.")
        if any(word in question_lower for word in ['total', 'sum', 'average', 'mean', 'max', 'min']):
            context_additions.append("The user wants statistical calculations - use the numeric summary data provided.")
        enhanced_context = "\n".join(context_additions) if context_additions else ""
        history_context = ""
        if qa_history:
            for qa in qa_history[-3:]:
                history_context += f"Q: {qa['question']}\nA: {qa['answer']}\n"
        prompt = f"""{data_summary}
{enhanced_context}
{history_context}
User's Current Question: {user_question}

Instructions:
1. Provide a clear, specific answer based on the data above
2. Consider the conversation history when relevant
3. If data quality issues exist, mention how they might affect your analysis
4. Use actual numbers from the data when possible
5. If the question cannot be fully answered, explain what information is available
6. For trend questions, reference the trend analysis
7. For financial data, consider business implications
8. Keep your response professional but accessible

Answer:"""

        chat_completion = client.chat.completions.create(
            model=model,
            messages=[
                {"role": "system", "content": "You are a professional financial data analyst with expertise in Excel analysis, trend identification, and business intelligence. You maintain context from previous questions about the same dataset."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.3,
            max_tokens=600,
        )
        return chat_completion.choices[0].message.content.strip()
    except Exception as e:
        return f"Error analyzing data: {str(e)}. Please check your OpenAI API configuration."

def store_file_data(df, file_info, errors, trends, data_summary):
    """Store dataframe and metadata in memory under a unique identifier."""
    file_id = str(uuid.uuid4())
    file_store[file_id] = {
        'dataframe': df,
        'file_info': file_info,
        'errors': errors,
        'trends': trends,
        'data_summary': data_summary,
        'qa_history': [],
        'timestamp': datetime.now()
    }
    return file_id

def get_file_data(file_id):
    """Retrieve stored file data."""
    return file_store.get(file_id)

def add_qa_to_history(file_id, question, answer):
    """Add Q&A to conversation history."""
    if file_id in file_store:
        file_store[file_id]['qa_history'].append({
            'question': question,
            'answer': answer,
            'timestamp': datetime.now()
        })

@app.route('/')
def index():
    """Render the main page."""
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    """Process uploaded Excel file and send first question to OpenAI."""
    if 'file' not in request.files:
        flash('Please choose an Excel file to upload.')
        return redirect(url_for('index'))

    file = request.files['file']
    query = request.form.get('question', '').strip()

    if file.filename == '':
        flash('Please choose an Excel file to upload.')
        return redirect(url_for('index'))

    if not query:
        flash('Enter a question about your spreadsheet so the AI knows what to analyze.')
        return redirect(url_for('index'))

    if not allowed_file(file.filename):
        flash('Unsupported file format. Upload a .xls or .xlsx spreadsheet.')
        return redirect(url_for('index'))

    try:
        filename = secure_filename(file.filename)
        print(f"Processing: {filename} with question: {query}")
        file_extension = filename.rsplit('.', 1)[1].lower()
        engine = 'openpyxl' if file_extension == 'xlsx' else 'xlrd'
        df = pd.read_excel(file, engine=engine)

        numeric_stats = {
            col: {
                'sum': float(df[col].sum()),
                'mean': float(df[col].mean()),
                'count': int(df[col].count())
            }
            for col in df.select_dtypes(include=['number']).columns
        }

        try:
            xl_file = pd.ExcelFile(file, engine=engine)
            sheet_names = xl_file.sheet_names
        except:
            sheet_names = ['Sheet1']

        file_info = {
            'filename': filename,
            'sheet_names': sheet_names,
            'num_rows': len(df),
            'num_columns': len(df.columns),
            'column_names': list(df.columns),
            'numeric_stats': numeric_stats
        }

        data_summary, errors, trends = create_enhanced_data_summary(df, file_info, file)
        file_id = store_file_data(df, file_info, errors, trends, data_summary)

        print("Sending enhanced request to OpenAI...")
        ai_answer = ask_openai_with_enhanced_context(data_summary, query, errors, trends)
        print(f"AI Response received: {len(ai_answer)} characters")

        add_qa_to_history(file_id, query, ai_answer)

        display_df = df.head(100) if len(df) > 100 else df
        file_data = display_df.to_html(classes='table table-bordered table-striped', index=False, table_id='data-table')

        return render_template('chat.html',
                               file_data=file_data,
                               file_info=file_info,
                               errors=errors,
                               trends=trends,
                               file_id=file_id,
                               qa_history=file_store[file_id]['qa_history'])
    except Exception as e:
        print(f"Error processing file: {str(e)}")
        flash("We couldn't read your Excel file. Try uploading a .xlsx file with data in the first sheet.")
        return redirect(url_for('index'))

@app.route('/ask_question', methods=['POST'])
def ask_question():
    """Answer additional user questions using stored file data."""
    file_id = request.form.get('file_id')
    question = request.form.get('question', '').strip()

    if not file_id or not question:
        return jsonify({'error': 'Missing file ID or question'}), 400

    file_data = get_file_data(file_id)

def create_enhanced_data_summary(df, file_info, file):
    """Create a summary of the data, detect errors, and find trends."""
    # Basic summary
    summary = f"Spreadsheet '{file_info['filename']}' with {file_info['num_rows']} rows and {file_info['num_columns']} columns.\n"
    summary += f"Columns: {', '.join(file_info['column_names'])}\n"
    # Detect errors
    errors = detect_data_errors(df)
    # Find trends (simple example: check if numeric columns are increasing or decreasing)
    trends = []
    for col in df.select_dtypes(include=[np.number]).columns:
        col_data = df[col].dropna()
        if len(col_data) > 1:
            direction = "increasing" if col_data.iloc[-1] > col_data.iloc[0] else "decreasing"
            trends.append({'column': col, 'direction': direction})
    return summary, errors, trends

if __name__ == "__main__":
    app.run(debug=True)
