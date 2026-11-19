"""Azure Function (Python v2 model): POST a Beneficiary Nomination PDF, get its data as JSON.

POST /api/extract-nomination
  Header: x-functions-key: <function key>
  Body:   the PDF file (binary), or JSON {"$content": "<base64>"} as Power Automate sometimes sends it.
Returns 200 with the result of nomination_extract.extract(); "valid": false lists the problems in "errors".
"""
import base64
import json

import azure.functions as func

from nomination_extract import extract

app = func.FunctionApp(http_auth_level=func.AuthLevel.FUNCTION)

MAX_BYTES = 20 * 1024 * 1024


def _pdf_bytes(req: func.HttpRequest) -> bytes:
    body = req.get_body() or b""
    if body[:5] == b"%PDF-":
        return body
    # Power Automate file content object: {"$content-type": "...", "$content": "<base64>"}
    try:
        data = json.loads(body)
        return base64.b64decode(data["$content"])
    except Exception:
        return body  # let extract() report it as unreadable


@app.route(route="extract-nomination", methods=["POST"])
def extract_nomination(req: func.HttpRequest) -> func.HttpResponse:
    pdf = _pdf_bytes(req)
    if not pdf:
        return _json({"valid": False, "errors": ["No file was sent."]}, 400)
    if len(pdf) > MAX_BYTES:
        return _json({"valid": False, "errors": ["File is larger than 20 MB."]}, 413)
    return _json(extract(pdf), 200)


def _json(obj, status):
    return func.HttpResponse(json.dumps(obj), status_code=status, mimetype="application/json")
