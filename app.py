import os
import urllib.parse
from flask import Flask, redirect, request, session, url_for, jsonify
import requests
import msal

app = Flask(__name__)

# Mandatory for Flask session cookies
app.secret_key = os.environ.get("SECRET_KEY", "change-me")

CLIENT_ID = os.environ.get("CLIENT_ID")
CLIENT_SECRET = os.environ.get("CLIENT_SECRET")
TENANT_ID = os.environ.get("TENANT_ID")  # ok sa fie tenantul tau, dar folosim authority "common"
BASE_URL = os.environ.get("BASE_URL", "https://personal-agent-zymv.onrender.com").rstrip("/")

TARGET_FOLDER = os.environ.get("TARGET_FOLDER", "Agent")
TARGET_FILE = os.environ.get("TARGET_FILE", "Master_Dorin_Agent.xlsx")

# For personal + work accounts
AUTHORITY = "https://login.microsoftonline.com/common"
REDIRECT_PATH = "/callback"
REDIRECT_URI = BASE_URL + REDIRECT_PATH

SCOPES = ["User.Read", "Files.ReadWrite", "offline_access"]


def msal_app(cache=None):
    return msal.ConfidentialClientApplication(
        CLIENT_ID,
        authority=AUTHORITY,
        client_credential=CLIENT_SECRET,
        token_cache=cache,
    )


def build_auth_url():
    return msal_app().get_authorization_request_url(
        scopes=SCOPES,
        redirect_uri=REDIRECT_URI,
        prompt="select_account",
    )


def get_token_from_code(code: str):
    result = msal_app().acquire_token_by_authorization_code(
        code,
        scopes=SCOPES,
        redirect_uri=REDIRECT_URI,
    )
    return result


def graph_get(url, access_token, params=None):
    headers = {"Authorization": f"Bearer {access_token}"}
    r = requests.get(url, headers=headers, params=params, timeout=30)
    return r


@app.route("/")
def home():
    return "OK"


@app.route("/ping")
def ping():
    return "pong"


@app.route("/login")
def login():
    # start Microsoft OAuth
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

    # store tokens in session (prototype)
    session["access_token"] = token_result["access_token"]
    if "refresh_token" in token_result:
        session["refresh_token"] = token_result["refresh_token"]

    return redirect(url_for("files"))


@app.route("/me")
def me():
    access_token = session.get("access_token")
    if not access_token:
        return redirect(url_for("login"))

    r = graph_get("https://graph.microsoft.com/v1.0/me", access_token)
    return (r.text, r.status_code, {"Content-Type": "application/json"})


def find_child_by_name(children, name):
    for c in children:
        if c.get("name", "").lower() == name.lower():
            return c
    return None


@app.route("/files")
def files():
    access_token = session.get("access_token")
    if not access_token:
        return redirect(url_for("login"))

    # list root children
    r = graph_get("https://graph.microsoft.com/v1.0/me/drive/root/children", access_token)
    if r.status_code != 200:
        return (r.text, r.status_code, {"Content-Type": "application/json"})

    root_children = r.json().get("value", [])
    folder = find_child_by_name(root_children, TARGET_FOLDER)
    if not folder:
        return jsonify({
            "ok": False,
            "msg": f"Folder not found in OneDrive root: {TARGET_FOLDER}",
            "hint": "Move folder to OneDrive root OR set TARGET_FOLDER correctly.",
            "root_items": [x.get("name") for x in root_children][:50],
        }), 404

    folder_id = folder["id"]

    # list folder children
    r2 = graph_get(f"https://graph.microsoft.com/v1.0/me/drive/items/{folder_id}/children", access_token)
    if r2.status_code != 200:
        return (r2.text, r2.status_code, {"Content-Type": "application/json"})

    items = r2.json().get("value", [])
    file_item = find_child_by_name(items, TARGET_FILE)

    return jsonify({
        "ok": True,
        "folder": TARGET_FOLDER,
        "file_expected": TARGET_FILE,
        "file_found": bool(file_item),
        "items": [{"name": x.get("name"), "id": x.get("id"), "type": ("folder" if "folder" in x else "file")} for x in items],
        "file_id": (file_item.get("id") if file_item else None),
        "next": "If file_found=true, we can read Excel cells next.",
    })


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=int(os.environ.get("PORT", "10000")))
