import os
import datetime
import pandas as pd
from flask import Flask, request, render_template_string

app = Flask(__name__)
FILE_PATH = "notes.xlsx"

def save_note(note):
    """Saves a note into an Excel spreadsheet."""
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    # Create a DataFrame for the new note
    new_data = pd.DataFrame([[timestamp, note]], columns=["Timestamp", "Note"])
    
    # Check if the file exists and append or create new
    if os.path.exists(FILE_PATH):
        existing_data = pd.read_excel(FILE_PATH)
        updated_data = pd.concat([existing_data, new_data], ignore_index=True)
    else:
        updated_data = new_data
    
    # Save to Excel
    updated_data.to_excel(FILE_PATH, index=False)
    return "Note saved successfully!"

@app.route("/", methods=["GET", "POST"])
def index():
    message = ""
    if request.method == "POST":
        note = request.form.get("note")
        if note:
            message = save_note(note)
    
    return render_template_string('''
        <html>
        <body>
            <h1>Submit a Note</h1>
            <form method="post">
                <textarea name="note" rows="4" cols="50"></textarea><br>
                <input type="submit" value="Save Note">
            </form>
            <p>{{ message }}</p>
        </body>
        </html>
    ''', message=message)

if __name__ == "__main__":
    app.run(debug=True)

