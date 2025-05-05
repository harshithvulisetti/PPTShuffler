import comtypes.client
import comtypes
import time
import win32clipboard
import logging
import pandas as pd
import os
from flask import Flask, request, render_template, send_file, flash, redirect, url_for
from werkzeug.utils import secure_filename
from werkzeug.exceptions import RequestEntityTooLarge
import uuid
import shutil

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

app = Flask(__name__)
app.secret_key = 'supersecretkey'
UPLOAD_FOLDER = 'Uploads'
ALLOWED_EXTENSIONS = {'xlsx', 'pptx'}
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 1000 * 1024 * 1024  # 50MB limit

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def get_shuffle_order(file1_path, file2_path, column_name="Member"):
    try:
        if not os.path.exists(file1_path):
            return [], f"Error: {file1_path} does not exist."
        if not os.path.exists(file2_path):
            return [], f"Error: {file2_path} does not exist."

        logging.info(f"Reading {file1_path}...")
        df1 = pd.read_excel(file1_path)
        logging.info(f"Reading {file2_path}...")
        df2 = pd.read_excel(file2_path)

        if column_name not in df1.columns:
            return [], f"Error: Column '{column_name}' not found in {file1_path}."
        if column_name not in df2.columns:
            return [], f"Error: Column '{column_name}' not found in {file2_path}."

        members1 = df1[column_name].tolist()
        members2 = df2[column_name].tolist()

        if len(members1) == 0 or len(members2) == 0:
            return [], "Error: One or both member lists are empty."
        if len(members1) != len(members2):
            return [], f"Error: Member counts differ ({len(members1)} in {file1_path}, {len(members2)} in {file2_path})."
        if sorted(members1) != sorted(members2):
            return [], "Error: Member names are not identical between the two files."

        shuffle_order = []
        for member in members2:
            index = members1.index(member) + 1
            shuffle_order.append(index)

        logging.info(f"Shuffle order: {shuffle_order}")
        return shuffle_order, None

    except Exception as e:
        logging.error(f"Error processing files: {e}")
        return [], f"Error processing files: {e}"

def clear_clipboard():
    try:
        win32clipboard.OpenClipboard()
        win32clipboard.EmptyClipboard()
        win32clipboard.CloseClipboard()
    except Exception as e:
        logging.warning(f"Failed to clear clipboard: {e}")

def shuffle_slides(input_pptx, output_pptx, slide_order):
    # Initialize COM for the current thread
    try:
        comtypes.CoInitialize()
    except Exception as e:
        logging.error(f"Error initializing COM: {e}")
        return False, f"Error initializing COM: {e}"

    logging.info("Initializing PowerPoint...")
    try:
        powerpoint = comtypes.client.CreateObject("PowerPoint.Application")
        powerpoint.Visible = True
    except Exception as e:
        logging.error(f"Error initializing PowerPoint: {e}")
        comtypes.CoUninitialize()
        return False, f"Error initializing PowerPoint: {e}"

    try:
        logging.info(f"Opening presentation: {input_pptx}")
        presentation = powerpoint.Presentations.Open(os.path.abspath(input_pptx))
        slides = presentation.Slides
        slide_count = slides.Count
        logging.info(f"Found {slide_count} slides")

        if slide_count == 0:
            presentation.Close()
            powerpoint.Quit()
            comtypes.CoUninitialize()
            return False, "No slides found in the presentation."

        if not slide_order:
            presentation.Close()
            powerpoint.Quit()
            comtypes.CoUninitialize()
            return False, "Slide order list is empty."

        if len(slide_order) != slide_count:
            presentation.Close()
            powerpoint.Quit()
            comtypes.CoUninitialize()
            return False, f"Slide order list length ({len(slide_order)}) does not match slide count ({slide_count})."

        if sorted(slide_order) != list(range(1, slide_count + 1)):
            presentation.Close()
            powerpoint.Quit()
            comtypes.CoUninitialize()
            return False, f"Slide order {slide_order} is invalid. Must contain unique indices from 1 to {slide_count}."

        logging.info(f"Using slide order: {slide_order}")
        new_presentation = powerpoint.Presentations.Add()

        for index in slide_order:
            logging.info(f"Copying slide {index}...")
            try:
                clear_clipboard()
                slides(index).Copy()
                time.sleep(0.5)
                new_presentation.Slides.Paste()
                time.sleep(0.5)
            except Exception as e:
                logging.error(f"Failed to copy/paste slide {index}: {e}")
                continue

        logging.info(f"Saving shuffled presentation as {output_pptx}")
        new_presentation.SaveAs(os.path.abspath(output_pptx))
        presentation.Close()
        new_presentation.Close()
        powerpoint.Quit()
        comtypes.CoUninitialize()
        return True, None

    except Exception as e:
        logging.error(f"Error during slide shuffling: {e}")
        try:
            presentation.Close()
        except:
            pass
        try:
            new_presentation.Close()
        except:
            pass
        try:
            powerpoint.Quit()
        except:
            pass
        comtypes.CoUninitialize()
        return False, f"Error during slide shuffling: {e}"
    finally:
        try:
            comtypes.CoUninitialize()
        except:
            pass

@app.errorhandler(RequestEntityTooLarge)
def handle_file_too_large(e):
    flash('Uploaded files are too large. Total size must be under 50MB.', 'error')
    return redirect(url_for('index'))

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/process', methods=['POST'])
def process():
    if not os.path.exists(UPLOAD_FOLDER):
        os.makedirs(UPLOAD_FOLDER)

    file1 = request.files.get('file1')
    file2 = request.files.get('file2')
    pptx_file = request.files.get('pptx_file')
    column_name = request.form.get('column_name', 'Member')

    if not file1 or not file2 or not pptx_file:
        flash('All files are required.', 'error')
        return redirect(url_for('index'))

    if not (allowed_file(file1.filename) and allowed_file(file2.filename) and allowed_file(pptx_file.filename)):
        flash('Invalid file type. Only .xlsx and .pptx files are allowed.', 'error')
        return redirect(url_for('index'))

    file1_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(file1.filename))
    file2_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(file2.filename))
    pptx_path = os.path.join(app.config['UPLOAD_FOLDER'], secure_filename(pptx_file.filename))
    output_path = os.path.join(app.config['UPLOAD_FOLDER'], f"shuffled_{uuid.uuid4().hex}.pptx")

    try:
        file1.save(file1_path)
        file2.save(file2_path)
        pptx_file.save(pptx_path)

        shuffle_order, error = get_shuffle_order(file1_path, file2_path, column_name)
        if error:
            flash(error, 'error')
            return redirect(url_for('index'))

        success, error = shuffle_slides(pptx_path, output_path, shuffle_order)
        if not success:
            flash(error, 'error')
            return redirect(url_for('index'))

        return render_template('result.html', output_file=os.path.basename(output_path))

    except Exception as e:
        flash(f"Error processing files: {e}", 'error')
        return redirect(url_for('index'))
    finally:
        for path in [file1_path, file2_path, pptx_path]:
            if os.path.exists(path):
                os.remove(path)

@app.route('/download/<filename>')
def download(filename):
    file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
    if os.path.exists(file_path):
        return send_file(file_path, as_attachment=True)
    flash('File not found.', 'error')
    return redirect(url_for('index'))

if __name__ == '__main__':
    app.run(debug=True)