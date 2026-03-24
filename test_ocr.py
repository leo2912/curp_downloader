import requests
import base64

def ocr_space_file(filename, overlay=False, api_key='helloworld', language='spa'):
    payload = {'isOverlayRequired': overlay,
               'apikey': api_key,
               'language': language,
               'OCREngine': 1}
    with open(filename, 'rb') as f:
        r = requests.post('https://api.ocr.space/parse/image',
                          files={filename: f},
                          data=payload)
    return r.json()

res = ocr_space_file('/home/deck/Documents/GitPrograms/curp_extractor/curp_downloader/test_ocr.py')
print(res)
