from flask import Flask, render_template, request, flash, redirect, url_for
import os
import pandas as pd
import numpy as np
from werkzeug.utils import secure_filename
from openai import OpenAI
from dotenv import load_dotenv
import openpyxl
from datetime import datetime
import re

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


def detect_data_errors(df):
    """Detect common data errors in the spreadsheet"""
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
            missing_percentage = (missing_count / len(df)) * 100
            if missing_percentage > 10:
                errors.append(f"Column '{col}' has {missing_count} missing values ({missing_percentage:.1f}%)")

    for col in numeric_cols:
        if len(df[col].dropna()) > 0:
            Q1 = df[col].quantile(0.25)
            Q3 = df[col].quantile(0.75)
            IQR = Q3 - Q1
            lower_bound = Q1 - 1.5 * IQR
            upper_bound = Q3 + 1.5 * IQR
            outliers = df[(df[col] < lower_bound) | (df[col] > upper_bound)][col]
            if len(outliers) > 0:
                errors.append(f"Column '{col}' has {len(outliers)} potential outliers (values significantly outside normal range)")
    
    return errors


def analyze_time_series_trends(df):
    """Analyze trends in time-series data"""
    trends = []
    

    date_columns = []
    for col in df.columns:
        if df[col].dtype == 'datetime64[ns]' or 'date' in col.lower() or 'time' in col.lower():
            date_columns.append(col)
    

    if not date_columns and len(df.columns) > 0:
        first_col = df.columns[0]
        try:
            pd.to_datetime(df[first_col].head(5), errors='raise')
            date_columns.append(first_col)
        except:
            pass
    
    if date_columns:
        date_col = date_columns[0]
     
        if df[date_col].dtype != 'datetime64[ns]':
            try:
                df[date_col] = pd.to_datetime(df[date_col])
            except:
                return trends
        
     
        numeric_cols = df.select_dtypes(include=[np.number]).columns
        for col in numeric_cols:
            if len(df[col].dropna()) >= 3:  
                # Sort by date
                df_sorted = df.sort_values(date_col)
                values = df_sorted[col].dropna()
                
                if len(values) >= 3:
                 
                    first_val = values.iloc[0]
                    last_val = values.iloc[-1]
                    change = last_val - first_val
                    change_percent = (change / first_val * 100) if first_val != 0 else 0
                    
                 
                    max_val = values.max()
                    min_val = values.min()
                    
                    trend_direction = "increasing" if change > 0 else "decreasing" if change < 0 else "stable"
                    
                    trends.append({
                        'column': col,
                        'direction': trend_direction,
                        'start_value': first_val,
                        'end_value': last_val,
                        'change': change,
                        'change_percent': change_percent,
                        'peak': max_val,
                        'minimum': min_val
                    })
    
    return trends


def extract_formulas_from_excel(file_path, sheet_name=None):
    """Extract formulas from Excel file using openpyxl"""
    formulas = []
    try:
        wb = openpyxl.load_workbook(file_path, data_only=False)
        ws = wb.active if sheet_name is None else wb[sheet_name]
        
        for row in ws.iter_rows():
            for cell in row:
                if cell.data_type == 'f' and cell.value: 
                    formulas.append({
                        'cell': f"{cell.column_letter}{cell.row}",
                        'formula': cell.value
                    })
        
        wb.close()
    except Exception as e:
        print(f"Error extracting formulas: {e}")
    
    return formulas


def create_enhanced_data_summary(df, file_info, file_obj=None):
    """Create an enhanced summary with error detection and trend analysis"""
    

    summary = f"""
EXCEL DATA ANALYSIS REPORT
==========================
File: {file_info['filename']}
Dimensions: {file_info['rows']:,} rows × {file_info['columns']} columns
Sheets: {', '.join(file_info['sheet_names'])}
Columns: {', '.join(file_info['column_names'])}
"""
    

    errors = detect_data_errors(df)
    if errors:
        summary += f"\n⚠️  DATA QUALITY ISSUES DETECTED:\n"
        for error in errors[:5]: 
            summary += f"- {error}\n"
        if len(errors) > 5:
            summary += f"... and {len(errors) - 5} more issues\n"
    else:
        summary += f"\n✅ No major data quality issues detected\n"
    
    trends = analyze_time_series_trends(df.copy())
    if trends:
        summary += f"\n📈 TREND ANALYSIS:\n"
        for trend in trends[:3]:
            summary += f"- {trend['column']}: {trend['direction'].upper()} trend "
            summary += f"({trend['start_value']:.2f} → {trend['end_value']:.2f}, "
            summary += f"{trend['change_percent']:+.1f}%)\n"
            summary += f"  Peak: {trend['peak']:.2f}, Minimum: {trend['minimum']:.2f}\n"
    if file_info.get('numeric_stats'):
        summary += f"\n📊 NUMERIC SUMMARY:\n"
        for col, stats in list(file_info['numeric_stats'].items())[:5]:  # Limit columns
            summary += f"- {col}: Sum={stats['sum']:,.2f}, Avg={stats['mean']:.2f}, Count={stats['count']}\n"
    
    if len(df) <= 10:
        summary += f"\n📋 COMPLETE DATASET:\n{df.to_string()}\n"
    else:
        summary += f"\n📋 SAMPLE DATA (first 5 rows):\n{df.head(5).to_string()}\n"
    
    return summary, errors, trends


def ask_openai_with_enhanced_context(data_summary, user_question, errors, trends, model="gpt-3.5-turbo"):
    """Enhanced AI query with comprehensive financial analysis context"""
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
        
        prompt = f"""You are an expert financial data analyst. Analyze the Excel data below and answer the user's question with professional insights.

{data_summary}

ANALYSIS GUIDANCE:
{enhanced_context}

User's Question: {user_question}

Instructions:
1. Provide a clear, specific answer based on the data above
2. If data quality issues exist, mention how they might affect your analysis
3. Use actual numbers from the data when possible
4. If the question cannot be fully answered, explain what information is available
5. For trend questions, reference the trend analysis
6. For financial data, consider business implications
7. Keep your response professional but accessible

Answer:"""

        chat_completion = client.chat.completions.create(
            model=model,
            messages=[
                {"role": "system", "content": "You are a professional financial data analyst with expertise in Excel analysis, trend identification, and business intelligence."},
                {"role": "user", "content": prompt}
            ],
            temperature=0.3,
            max_tokens=600,
        )
        return chat_completion.choices[0].message.content.strip()
    except Exception as e:
        return f"Error analyzing data: {str(e)}. Please check your OpenAI API configuration."


@app.route('/')
def index():
    """Render the main page"""
    return render_template('index.html')


@app.route('/upload', methods=['GET', 'POST'])
def upload_file():
    """Handle file upload and enhanced question processing"""
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
        print(f"Processing: {filename} with question: {query}")
        
        file_extension = filename.rsplit('.', 1)[1].lower()
        engine = 'openpyxl' if file_extension == 'xlsx' else 'xlrd'
        
        df = pd.read_excel(file, engine=engine)
        
        numeric_stats = {}
        for col in df.select_dtypes(include=['number']).columns:
            numeric_stats[col] = {
                'sum': float(df[col].sum()),
                'mean': float(df[col].mean()),
                'count': int(df[col].count())
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
        
        data_summary, errors, trends = create_enhanced_data_summary(df, file_info, file)
        
        print("Sending enhanced request to OpenAI...")
        ai_answer = ask_openai_with_enhanced_context(data_summary, query, errors, trends)
        print(f"AI Response received: {len(ai_answer)} characters")
        
        display_df = df.head(100) if len(df) > 100 else df
        file_data = display_df.to_html(classes='table table-bordered table-striped', 
                                      index=False, 
                                      table_id='data-table')
        
        return render_template('index.html', 
                             file_data=file_data, 
                             query=query, 
                             file_info=file_info,
                             ai_answer=ai_answer,
                             errors=errors,
                             trends=trends,
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
