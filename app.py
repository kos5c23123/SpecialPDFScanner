from flask import Flask, render_template, request, send_file
import os
import pandas as pd
import re
import threading
from concurrent.futures import ThreadPoolExecutor
from PIL import Image
import pytesseract
import pypdfium2
from io import BytesIO

app = Flask(__name__)
lock = threading.Lock()


def clear_uploads_folder():
    folder = 'uploads'
    for filename in os.listdir(folder):
        file_path = os.path.join(folder, filename)
        try:
            if os.path.isfile(file_path) or os.path.islink(file_path):
                os.unlink(file_path)  # Remove the file
            elif os.path.isdir(file_path):
                shutil.rmtree(file_path)  # Remove the directory
        except Exception as e:
            print(f'Failed to delete {file_path}. Reason: {e}')


def combine_next_number(data):
    result = []
    i = 0
    while i < len(data):
        if data[i] == '1' and i + 1 < len(data):
            # Combine '1' with the next number
            combined = data[i] + data[i + 1]
            result.append(combined)
            i += 2  # Skip the next number since it's already combined
        else:
            # Just add the current element to the result
            result.append(data[i])
            i += 1
    return result

def extract_relevant_number(numbers):
    """
    Extracts the number based on the criteria:
    - The number should be between 600 and 1500
    - The previous number must be less than 100
    - The number should not be in the skipped values

    Parameters:
    numbers (list of str): List of numbers as strings.
    skipped_values (set of float): Set of numbers to skip.

    Returns:
    float: The relevant number or None if no valid number is found.
    """
    # Convert skipped values to float for comparison
    skipped_values = {'778', '683', '1123', '1581'}
    skipped_set = set(map(float, skipped_values))

    # Convert numbers to floats and filter based on the range and skipped values
    filtered_numbers = []
    for num in numbers:
        if re.match(r'^\d+(\.\d+)?$', num):
            num_float = float(num)
            if 600 <= num_float <= 1400 and num_float not in skipped_set:
                filtered_numbers.append(num_float)

    print(f"Filtered Numbers: {filtered_numbers}")

    if filtered_numbers:
        # return max(filtered_numbers)
        return filtered_numbers[0]
    return None

def extract_number_from_specific_area(pdf_path):
    with lock:
        pdf = pypdfium2.PdfDocument(pdf_path)
        page = pdf[0]

        scale = 6.0
        pil_image = page.render(scale=scale).to_pil()

        left, top, width, height = 0, 2200, 1200, 900
        # left, top, width, height = 0, 1900, 800, 600
        box = (left, top, left + width, top + height)
        cropped_image = pil_image.crop(box)

        cropped_image.save("output_image.png")

        text = pytesseract.image_to_string(cropped_image, config='--psm 6')
        numbers = re.findall(r'\b\d+(?:\.\d+)?\b', text)

        new_num = combine_next_number(numbers)

        result = extract_relevant_number(new_num)

        return [result, new_num]

def process_file(file):
    pdf_path = os.path.join("uploads", file.filename)
    file.save(pdf_path)

    result = extract_number_from_specific_area(pdf_path)

    return {"Filename": os.path.splitext(file.filename)[0], "Extracted Number": result[0]}
    # return {"Filename": os.path.splitext(file.filename)[0], "Extracted Number": result[0], "Number List": result[1]}

@app.route('/', methods=['GET', 'POST'])
def index():
    if request.method == 'POST':
        # Clear the uploads folder before saving new files
        clear_uploads_folder()
        
        files = request.files.getlist('pdf_files')
        results = []

        with ThreadPoolExecutor() as executor:
            results = list(executor.map(process_file, files))

        df = pd.DataFrame(results)
        excel_path = os.path.join("uploads", "ocr_results.xlsx")
        df.to_excel(excel_path, index=False)

        return render_template('result.html', results=results, excel_file='ocr_results.xlsx')

    return render_template('index.html')


@app.route('/download/<filename>')
def download_file(filename):
    file_path = os.path.join('uploads', filename)
    return send_file(file_path, as_attachment=True)

if __name__ == '__main__':
    if not os.path.exists('uploads'):
        os.makedirs('uploads')
    app.run(host='0.0.0.0', port=5001, debug=True)
