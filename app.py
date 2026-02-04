from flask import Flask, request, Response
import html

app = Flask(__name__)

def twiml(message: str) -> Response:
    safe = html.escape(message)
    xml = f'<?xml version="1.0" encoding="UTF-8"?><Response><Message>{safe}</Message></Response>'
    return Response(xml, mimetype="text/xml")

@app.get("/")
def health():
    return "OK", 200

@app.post("/twilio/inbound")
def inbound():
    body = (request.form.get("Body") or "").strip()
    body_l = body.lower()

    if body_l in ("ping", "test"):
        return twiml("pong")

    if body_l in ("status", "stare"):
        return twiml(
            "status:\n"
            "- webhook: OK\n"
            "- mode: sandbox\n"
            "Comenzi: ping, status, help"
        )

    if body_l in ("help", "ajutor", "?"):
        return twiml(
            "Comenzi disponibile:\n"
            "- ping\n"
            "- status\n"  
            "- help\n"
            "\n"
            "Urmeaza: conectam Excel + alerte."
        )

    return twiml("Am primit: " + body + "\nScrie 'help' pentru comenzi.")
