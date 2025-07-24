import os
from flask import Flask, request, render_template, redirect, flash
from werkzeug.utils import secure_filename
from datetime import datetime
from automation import process_onboarding

app = Flask(__name__)
app.secret_key = '1234567'


app.config['UPLOAD_FOLDER'] = os.path.abspath(os.path.join(os.path.dirname(__file__), '..', 'data'))

ALLOWED_EXTENSIONS = {'xlsx'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        if 'excel_file' not in request.files:
            flash('No file part')
            return redirect(request.url)
        file = request.files['excel_file']
        if file.filename == '':
            flash('No file selected')
            return redirect(request.url)
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            timestamp = datetime.now().strftime("%Y%m%d%H%M%S")
            unique_filename = f"{timestamp}_{filename}"
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], unique_filename)

            try:
                file.save(filepath)
            except PermissionError as e:
                flash(f"Permission denied when saving the file: {e}")
                return redirect(request.url)

            sender_email = request.form.get('sender_email')
            sender_password = request.form.get('sender_password')

            if not sender_email or not sender_password:
                flash("Please enter both sender email and password")
                return redirect(request.url)

            try:
                sent, failed = process_onboarding(
                    excel_path=filepath,
                    sender_email=sender_email,
                    sender_password=sender_password,
                    email_host='smtp.gmail.com',
                    email_port=587,
                    smtp_ssl=False,
                    smtp_starttls=True
                )
            except Exception as e:
                flash(f"Error during automation: {e}")
                return redirect(request.url)

            return render_template('result.html', sent=sent, failed=failed)
        else:
            flash("Allowed file types: xlsx")
            return redirect(request.url)

    return render_template('index.html')


if __name__ == '__main__':
    app.run(debug=True)
