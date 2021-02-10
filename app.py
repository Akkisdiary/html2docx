from flask import Flask
from .htmldocx import HtmlToDocx


app = Flask(__name__)
new_parser = HtmlToDocx()
file = './my_tests/biden.html'


@app.route('/')
def hello_world():
    new_parser.parse_html_file(file)
    return 'Hello, World!'


if __name__ == "__main__":
    app.run(debug=True)
