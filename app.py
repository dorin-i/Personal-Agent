from flask import Flask, request, redirect, url_for

app = Flask(__name__)


@app.route("/")
def home():
    return """
    <h1>Personal Agent</h1>
    <p>Server running OK.</p>
    <a href="/login">Login</a>
    """


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        user = request.form.get("username")
        return f"<h2>Welcome, {user}</h2>"

    return """
    <h2>Login</h2>
    <form method="post">
        <input name="username" placeholder="User"><br><br>
        <button type="submit">Login</button>
    </form>
    """


if __name__ == "__main__":
    app.run(host="0.0.0.0", port=10000)
