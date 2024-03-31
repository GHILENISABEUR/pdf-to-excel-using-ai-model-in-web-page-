from flask import Flask, render_template, request, redirect, url_for, send_from_directory
import os
import PyPDF2
import pandas as pd
import fitz  # PyMuPDF
import xlsxwriter
import base64
import imghdr

app = Flask(__name__)

UPLOAD_FOLDER = 'uploads'
ALLOWED_EXTENSIONS = {'pdf',}

app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

def allowed_file(filename):
    return '.' in filename and \
           filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS
def pdf_to_excel(pdf_path, excel_path):
    # Open the PDF file
    pdf_document = fitz.open(pdf_path)
    num_pages = pdf_document.page_count
    # Initialize empty lists to store the text and images from each page
    text_data = []
    image_data = []
    # Extract text and images from each page
    for page_num in range(num_pages):
        page = pdf_document.load_page(page_num)
        text_data.append(page.get_text())
        images = page.get_images(full=True)
        for img_index, img_info in enumerate(images):
            img_bytes = pdf_document.extract_image(img_info[0])
            image_data.append(img_bytes["image"])
    # Convert the text data into a DataFrame
    text_df = pd.DataFrame(text_data, columns=['Text'])
    # Convert the image data into a DataFrame
    image_df = pd.DataFrame(image_data, columns=['Image'])
    # Write DataFrames to Excel
    with pd.ExcelWriter(excel_path, engine='xlsxwriter') as writer:
        text_df.to_excel(writer, sheet_name='Text', index=False)
        # Create a separate sheet for images
        workbook = writer.book
        worksheet = workbook.add_worksheet('Images')
        for row_index, image_bytes in enumerate(image_data):
            img_base64 = base64.b64encode(image_bytes).decode('utf-8')
            img_format = imghdr.what(None, image_bytes)
            image_path = f'image_{row_index + 1}.{img_format}'
            with open(image_path, 'wb') as img_file:
                img_file.write(image_bytes)
            worksheet.insert_image(row_index, 0, image_path, {'x_scale': 0.5, 'y_scale': 0.5})

    # Close the PDF document
    pdf_document.close()

@app.route('/', methods=['GET', 'POST'])
def upload_file():
    if request.method == 'POST':
        # Check if the post request has the file part
        if 'file' not in request.files:
            return redirect(request.url)
        file = request.files['file']
        # If the user does not select a file, the browser submits an empty file without a filename
        if file.filename == '':
            return redirect(request.url)
        if file and allowed_file(file.filename):
            filename = file.filename
            # Create the uploads directory if it doesn't exist
            if not os.path.exists(app.config['UPLOAD_FOLDER']):
                os.makedirs(app.config['UPLOAD_FOLDER'])
            # Save the uploaded file
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
            # Generate Excel filename
            excel_filename = os.path.splitext(filename)[0] + '.xlsx'
            excel_path = os.path.join(app.config['UPLOAD_FOLDER'], excel_filename)
            # Convert PDF to Excel
            pdf_to_excel(os.path.join(app.config['UPLOAD_FOLDER'], filename), excel_path)
            # Redirect to the Excel file
            return redirect(url_for('uploaded_file', filename=excel_filename))
    return render_template('upload.html')

@app.route('/uploads/<filename>')
def uploaded_file(filename):
    return send_from_directory(app.config['UPLOAD_FOLDER'], filename)

if __name__ == '__main__':
    app.run(debug=True)
