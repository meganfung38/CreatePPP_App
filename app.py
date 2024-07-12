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
        # no file upload in POST request
        if 'excel_file' not in request.files:
            return render_template('index_html', error='Include an Excel file to create a PPP')

        file = request.files['excel_file']  # retrieve Excel file from POST request
        sheet_name = request.form.get('sheet_name')  # retrieve text input from POST request

        # checking if uploaded file upload is empty
        if file.filename == '':
            return render_template('index.html', error='Select an Excel file to create a PPP')

        # process uploaded file
        if file:
            # temporarily save the file to a secure location (uploads folder)
            filename = secure_filename(file.filename)
            file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            file.save(file_path)

            # Generate PPP report
            if sheet_name:  # sheet name was provided
                progress_output, plans_output, problems_output = generatePPP.create_ppp(file_path, pg=sheet_name)
            else:  # sheet name was not provided
                progress_output, plans_output, problems_output = generatePPP.create_ppp(file_path)

            # remove file from secure location once PPP is generated
            os.remove(file_path)

            # Render PPP report or pass it to a new template
            return render_template('index.html',
                                   progress_output=progress_output,
                                   plans_output=plans_output,
                                   problems_output=problems_output)

    except Exception as e:  # error occurred
        return render_template('index.html', error=str(e))


if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    app.run(host='0.0.0.0', port=port)
