from flask import Flask

app = Flask(__name__)

@app.route("/")
def home():
    return """
    <h1>CRM Web</h1>
    <h2>Application is running successfully!</h2>
    <p>Welcome to our CRM system.</p>
    """

if __name__ == "__main__":
    app.run(host="0.0.0.0", port=8000)
