import requests
from config import Config
from decipher.beacon import api

HEADERS = {
    "x-apikey": Config.DECIPHER_API_KEY,  # Updated to use the correct header for the API key
    "Accept": "application/json"
}

def lookup_survey(survey_id: str) -> list:
    Config.validate()

    url = f"{Config.DECIPHER_BASE}/api/v1/rh/surveys/selfserve/2227/{survey_id}"
    r = requests.get(url, headers=HEADERS)
    if r.status_code == 404:
        return [{"error": f"Survey ID {survey_id} not found."}]
    r.raise_for_status()
    if r.headers.get('Content-Type') == 'application/json':
        return [r.json()]  # Wrap the response in a list
    else:
        return [{"error": f"Unexpected response format: {r.status_code} {r.reason}", "content": r.text}]

def fetch_survey_xml(survey_id):
    Config.validate()

    url = f"{Config.DECIPHER_BASE}/api/v1/surveys/selfserve/2227/{survey_id}/files/survey.xml"
    print(f"Fetching XML for survey_id: {survey_id}")  # Log the survey_id
    r = requests.get(url, headers={
        "x-apikey": Config.DECIPHER_API_KEY,  # Corrected header name
        "Accept": "application/xml"
    })
    print(f"Response status: {r.status_code}")  # Log the response status
    if r.status_code != 200:
        print(f"Error response: {r.text}")  # Log error response
    r.raise_for_status()
    print(f"Fetched XML: {r.text[:200]}...")  # Log the first 200 characters of the XML
    return r.text
