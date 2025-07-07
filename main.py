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
app.config["MAX_CONTENT_LENGTH"] = 5 * 1024 * 1024 * 1024  # 5 GB max upload
app.config["SECRET_KEY"] = "your-secret-key-here"

file_store = {}
client = OpenAI()

ALLOWED_EXTENSIONS = {"xls", "xlsx"}

def allowed_file(filename):
    """Check if the uploaded file has an allowed extension."""
    return "." in filename and filename.rsplit(".", 1)[1].lower() in ALLOWED_EXTENSIONS

def detect_data_errors(df):
    """Detect common data errors in the spreadsheet."""
    errors = []
    for col in df.columns:
        for idx, value in enumerate(df[col]):
            if pd.isna(value):
                continue
            if isinstance(value, str):
                excel_errors = [
                    "#DIV/0!",
                    "#N/A",
                    "#NAME?",
                    "#NULL!",
                    "#NUM!",
                    "#REF!",
                    "#VALUE!",
                ]
                if any(error in str(value) for error in excel_errors):
                    errors.append(f"Excel error \'{value}\' found in column \'{col}\' at row {idx + 2}")
    numeric_cols = df.select_dtypes(include=[np.number]).columns
    for col in numeric_cols:
        missing_count = df[col].isna().sum()
        if missing_count > 0:
            errors.append(f"Column \'{col}\' has {missing_count} missing values")
    return errors


stored_dataframe = None

def get_specific_data_for_question(question, df):
    """Extract only relevant data based on the question to minimize tokens."""
    question_lower = question.lower()

    if len(df) <= 30:
        return ""

    if any(word in question_lower for word in ["degree", "value", "cell", "row"]):
        import re

        numbers = re.findall(r"\\d+", question)
        if numbers:
            try:
                for num in numbers[:3]:
                    val = int(num)

                    for col in df.columns:
                        matches = df[df[col] == val]
                        if not matches.empty:
                            return f"Rows with {col}={val}:\\n{matches.to_string(index=True)}\\n"
            except:
                pass

    if any(word in question_lower for word in ["sum", "total"]):
        return ""  # Return empty string for sum/total as it will be handled by data_summary

    return ""

def ask_openai_with_enhanced_context(
    data_summary, user_question, errors, trends, qa_history=None, model="gpt-3.5-turbo"
):
    """Send the user\'s question and data summary to OpenAI with extra context."""
    try:
        specific_data = ""
        if stored_dataframe is not None:
            specific_data = get_specific_data_for_question(user_question, stored_dataframe)

        prompt = f"{data_summary}{specific_data}Q: {user_question}\\nA:"

        chat_completion = client.chat.completions.create(
            model=model,
            messages=[
                {"role": "system", "content": "Answer directly using provided data. One sentence maximum."},
                {"role": "user", "content": prompt},
            ],
            temperature=0,
            max_tokens=50,
        )
        return chat_completion.choices[0].message.content.strip()
    except Exception as e:
        return f"Error analyzing data: {str(e)}. Please check your OpenAI API configuration."

def store_file_data(df, file_info, errors, trends, data_summary):
    """Store dataframe and metadata in memory under a unique identifier."""
    file_id = str(uuid.uuid4())
    file_store[file_id] = {
        "dataframe": df,
        "file_info": file_info,
        "errors": errors,
        "trends": trends,
        "data_summary": data_summary,
        "qa_history": [],
        "timestamp": datetime.now(),
    }
    return file_id

def get_file_data(file_id):
    """Retrieve stored file data."""
    return file_store.get(file_id)

def add_qa_to_history(file_id, question, answer):
    """Add Q&A to conversation history."""
    if file_id in file_store:
        file_store[file_id]["qa_history"].append(
            {"question": question, "answer": answer, "timestamp": datetime.now()}
        )


@app.route("/")
def index():
    """Render the main page."""
    return render_template("index.html")


@app.route("/upload", methods=["POST"])
def upload_file():
    """Process uploaded Excel file and send first question to OpenAI."""
    if "file" not in request.files:
        flash("📄 Please select an Excel file to upload before proceeding.")
        return redirect(url_for("index"))

    file = request.files["file"]
    query = request.form.get("question", "").strip()

    if file.filename == "":
        flash("📄 No file selected. Please choose an Excel file (.xlsx or .xls) from your computer.")
        return redirect(url_for("index"))

    if not query:
        flash("❓ Please enter a question about your data so our AI knows what to analyze. For example: \"What are the total sales?\" or \"Show me trends in the data.\"")
        return redirect(url_for("index"))

    if not allowed_file(file.filename):
        flash("❌ File format not supported. Please upload an Excel file with .xlsx or .xls extension. Other formats like .csv or .txt are not currently supported.")
        return redirect(url_for("index"))

    try:
        filename = secure_filename(file.filename)
        print(f"Processing: {filename} with question: {query}")
        file_extension = filename.rsplit(".", 1)[1].lower()
        engine = "openpyxl" if file_extension == "xlsx" else "xlrd"
        df = pd.read_excel(file, engine=engine)

        numeric_stats = {
            col:
                {
                    "sum": float(df[col].sum()),
                    "mean": float(df[col].mean()),
                    "count": int(df[col].count()),
                }
            for col in df.select_dtypes(include=["number"]).columns
        }

        try:
            # No need to specify engine for ExcelFile, pandas handles it.
            xl_file = pd.ExcelFile(file)
            sheet_names = xl_file.sheet_names
        except:
            sheet_names = ["Sheet1"]

        file_info = {
            "filename": filename,
            "sheet_names": sheet_names,
            "num_rows": len(df),
            "num_columns": len(df.columns),
            "column_names": list(df.columns),
            "numeric_stats": numeric_stats,
        }

        data_summary, errors, trends = create_enhanced_data_summary(df, file_info, file)
        file_id = store_file_data(df, file_info, errors, trends, data_summary)

        print("Sending enhanced request to OpenAI...")
        ai_answer = ask_openai_with_enhanced_context(data_summary, query, errors, trends)
        print(f"AI Response received: {len(ai_answer)} characters")

        add_qa_to_history(file_id, query, ai_answer)

        display_df = df.head(100) if len(df) > 100 else df
        file_data = display_df.to_html(classes="table table-bordered table-striped", index=False, table_id="data-table")

        return render_template(
            "chat.html",
            file_data=file_data,
            file_info=file_info,
            errors=errors,
            trends=trends,
            file_id=file_id,
            qa_history=file_store[file_id]["qa_history"],
        )
    except Exception as e:
        print(f"Error processing file: {str(e)}")
        flash("We couldn\'t read your Excel file. Try uploading a .xlsx file with data in the first sheet.")
        return redirect(url_for("index"))


@app.route("/ask_question", methods=["POST"])
def ask_question():
    """Answer additional user questions using stored file data."""
    file_id = request.form.get("file_id")
    question = request.form.get("question", "").strip()

    if not file_id or not question:
        return jsonify({"error": "Missing file ID or question"}), 400

    file_data = get_file_data(file_id)

    if not file_data:
        return jsonify({"error": "File data not found. Please upload your file again."}), 404

    try:
        global stored_dataframe
        stored_dataframe = file_data["dataframe"]

        # Get AI response
        ai_answer = ask_openai_with_enhanced_context(
            file_data["data_summary"],
            question,
            file_data["errors"],
            file_data["trends"],
        )

        add_qa_to_history(file_id, question, ai_answer)

        return jsonify(
            {"question": question, "answer": ai_answer, "timestamp": datetime.now().isoformat()}
        )
    except Exception as e:
        return jsonify({"error": f"Error processing question: {str(e)}"}), 500


def create_enhanced_data_summary(df, file_info, file):
    """Create a minimal data summary to reduce token usage."""

    global stored_dataframe
    stored_dataframe = df

    summary = f"Data: {file_info['num_rows']} rows, {file_info['num_columns']} columns\n"
    summary += f"Columns: {', '.join(file_info['column_names']) }\n"

    if len(df) <= 30:
        summary += f"Complete data:\n{df.to_string(index=True)}\n"
    else:
        summary += f"Sample data (first 5):\n{df.head(5).to_string(index=True)}\n"
        summary += f"Sample data (last 5):\n{df.tail(5).to_string(index=True)}\n"

    numeric_cols = df.select_dtypes(include=[np.number]).columns
    if len(numeric_cols) > 0:
        sums = {col: df[col].sum() for col in numeric_cols}
        summary += f"Column sums: {sums}\n"

    errors = detect_data_errors(df)

    trends = []
    for col in numeric_cols:
        values = df[col].dropna().values
        if len(values) < 2:
            trend = "insufficient data"
        elif all(x < y for x, y in zip(values, values[1:])):
            trend = "increasing"
        elif all(x > y for x, y in zip(values, values[1:])):
            trend = "decreasing"
        else:
            trend = "mixed"
        trends.append(f"{col}: {trend}")

    return summary, errors, trends


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=5000, debug=True)
