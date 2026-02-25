import os
from flask import Flask, redirect, request, session, url_for, jsonify
import requests
import msal

app = Flask(__name__)

# Mandatory for Flask session cookies
app.secret_key = os.environ.get("SECRET_KEY", "change-me")

CLIENT_ID = os.environ.get("CLIENT_ID")
CLIENT_SECRET = os.environ.get("CLIENT_SECRET")

BASE_URL = os.environ.get("BASE_URL", "https://personal-agent-zymv.onrender.com").rstrip("/")

TARGET_FOLDER = os.environ.get("TARGET_FOLDER", "Agent")  # not used in AppFolder mode
TARGET_FILE = os.environ.get("TARGET_FILE", "Master_Dorin_Agent.xlsx")
APP_FILE_NAME = TARGET_FILE

# For personal + work accounts
AUTHORITY = "https://login.microsoftonline.com/common"
REDIRECT_PATH = "/callback"
REDIRECT_URI = BASE_URL + REDIRECT_PATH

# AppFolder-only scope
SCOPES = ["User.Read", "Files.ReadWrite.AppFolder"]


def msal_app():
    return msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=AUTHORITY,
        client_credential=CLIENT_SECRET,
    )


def build_auth_url():
    return msal_app().get_authorization_request_url(
        scopes=SCOPES,
        redirect_uri=REDIRECT_URI,
        prompt="select_account",
    )


def get_token_from_code(code: str):
    return msal_app().acquire_token_by_authorization_code(
        code,
        scopes=SCOPES,
        redirect_uri=REDIRECT_URI,
    )


def graph_get(url, access_token, params=None):
    headers = {"Authorization": f"Bearer {access_token}"}
    return requests.get(url, headers=headers, params=params, timeout=30)


def graph_put(url, access_token, data: bytes, content_type: str):
    headers = {
        "Authorization": f"Bearer {access_token}",
        "Content-Type": content_type,
    }
    return requests.put(url, headers=headers, data=data, timeout=60)


def find_child_by_name(children, name):
    for c in children:
        if c.get("name", "").lower() == name.lower():
            return c
    return None


@app.route("/")
def home():
    return """
    <h1>Personal Agent</h1>
    <p>Server running OK.</p>
    <ul>
      <li><a href="/login">Login</a></li>
      <li><a href="/files">List AppFolder files</a></li>
      <li><a href="/upload">Upload Excel to AppFolder</a></li>
    </ul>
    """


@app.route("/login")
def login():
    return redirect(build_auth_url())


@app.route(REDIRECT_PATH)
def authorized():
    if "error" in request.args:
        return jsonify({"error": request.args.get("error"), "desc": request.args.get("error_description")}), 400

    code = request.args.get("code")
    if not code:
        return "Missing code", 400

    token_result = get_token_from_code(code)
    if "access_token" not in token_result:
        return jsonify(token_result), 400

    session["access_token"] = token_result["access_token"]
    return redirect(url_for("files"))


@app.route("/files")
def files():
    access_token = session.get("access_token")
    if not access_token:
        return redirect(url_for("login"))

    # AppFolder root (approot)
    r = graph_get("https://graph.microsoft.com/v1.0/me/drive/special/approot/children", access_token)
    if r.status_code != 200:
        return (r.text, r.status_code, {"Content-Type": "application/json"})

    items = r.json().get("value", [])
    file_item = find_child_by_name(items, APP_FILE_NAME)

    return jsonify({
        "ok": True,
        "scope": "AppFolder",
        "file_expected": APP_FILE_NAME,
        "file_found": bool(file_item),
        "items": [{"name": x.get("name"), "id": x.get("id"), "type": ("folder" if "folder" in x else "file")} for x in items],
        "file_id": (file_item.get("id") if file_item else None),
    })


@app.route("/upload", methods=["GET", "POST"])
def upload():
    access_token = session.get("access_token")
    if not access_token:
        return redirect(url_for("login"))

    if request.method == "GET":
        return f"""
        <h2>Upload Excel to AppFolder</h2>
        <p>Target file name: <b>{APP_FILE_NAME}</b></p>
        <form method="POST" enctype="multipart/form-data">
            <input type="file" name="file" accept=".xlsx" required />
            <button type="submit">Upload</button>
        </form>
        """

    f = request.files.get("file")
    if not f or f.filename == "":
        return "No file selected", 400

    if not f.filename.lower().endswith(".xlsx"):
        return "Only .xlsx allowed", 400

    data = f.read()

    url = f"https://graph.microsoft.com/v1.0/me/drive/special/approot:/{APP_FILE_NAME}:/content"
    r = graph_put(
        url,
        access_token,
        data=data,
        content_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    )

    if r.status_code not in (200, 201):
        return f"Upload failed: {r.status_code} {r.text}", 500

    return "Upload OK. Open /files to verify.", 200

RUN_TOKEN = os.environ.get("RUN_TOKEN", "")

@app.route("/run", methods=["GET"])
def run_job():
    token = request.args.get("token", "")
    job = request.args.get("job", "manual")

    if not RUN_TOKEN or token != RUN_TOKEN:
        return jsonify({"ok": False, "error": "unauthorized"}), 401

    return jsonify({"ok": True, "job": job, "note": "run endpoint OK (engine next)"}), 200

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", "10000")))
