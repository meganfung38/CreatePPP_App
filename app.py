import os
import generatePPP
from flask import Flask, render_template, request
from werkzeug.utils import secure_filename

app = Flask(__name__)

# Configuration for file uploads
UPLOAD_FOLDER = os.path.join(os.getcwd(), 'uploads')  # Folder to store uploaded files
if not os.path.exists(UPLOAD_FOLDER):  # create directory for uploads if DNE
    os.makedirs(UPLOAD_FOLDER)

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER


# render interface for Excel upload and text input
@app.route('/')
def index():
    return render_template('index.html')


# create a PPP for Excel upload and text input
@app.route('/generatePPP', methods=['POST'])
def generate_ppp():
    try:
        # Check if the post request has the file part
        if 'excel_file' not in request.files:
            return render_template('index_html', error='Include an Excel file to create a PPP')

        file = request.files['excel_file']

        # If user does not select file, browser also submit an empty part without filename
        if file.filename == '':
            return render_template('index.html', error='No selected file')

        if file:
            # Save the file to a secure location
            filename = secure_filename(file.filename)
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(file_path)

            # Generate PPP report
            ppp_report = generatePPP.create_ppp(file_path)

            # Render PPP report or pass it to a new template
            return render_template('index.html', ppp_report=ppp_report)

    except Exception as e:
        return render_template('index.html', error=str(e))


if __name__ == '__main__':
    app.run(debug=True)
