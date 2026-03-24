import requests

def test_api():
    url = "https://www.gob.mx/curp/"
    # Just a basic request to see if we get blocked
    headers = {
        "User-Agent": "Mozilla/5.0 (Windows NT 10.0; Win64; x64)"
    }
    r = requests.get(url, headers=headers)
    print(r.status_code)

test_api()
