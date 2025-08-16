import requests
import pandas as pd

# === CONFIGURATION ===
API_KEY = 'K81117496188957'  # Replace this with your OCR.Space API key
IMAGE_PATH = './first_scr.png'  # e.g. 'first.jpg'
OUTPUT_EXCEL = 'output.xlsx'
OCR_API_URL = 'https://api.ocr.space/parse/image'


# === STEP 1: CALL OCR.SPACE API ===
def extract_text_from_image(image_path, api_key):
    """Send image to OCR.Space API and extract raw text"""
    with open(image_path, 'rb') as image_file:
        response = requests.post(
            OCR_API_URL,
            files={'filename': image_file},
            data={
                'apikey': api_key,
                'language': 'eng',
                'isTable': True,        # Enable table parsing
                'OCREngine': 2          # Use OCR engine 2 (better accuracy)
            }
        )
    return response.json()


# === STEP 2: CONVERT PARSED TEXT TO EXCEL ===
def text_to_excel(parsed_text, output_file):
    """
    Converts raw parsed table text to Excel.
    Assumes tab-delimited text (common with table parsing).
    """
    lines = parsed_text.strip().split('\n')
    rows = [line.split('\t') for line in lines]

    df = pd.DataFrame(rows)
    df.to_excel(output_file, index=False, header=False)
    print(f"[✓] Data exported to Excel: {output_file}")


# === STEP 3: RUN THE PROCESS ===
def main():
    result = extract_text_from_image(IMAGE_PATH, API_KEY)

    try:
        parsed_text = result['ParsedResults'][0]['ParsedText']
        text_to_excel(parsed_text, OUTPUT_EXCEL)
    except Exception as e:
        print(f"[✗] Error extracting text or saving Excel: {e}")


if __name__ == '__main__':
    main()
