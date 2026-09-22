import os
import sys

from flask import Flask, render_template

app = Flask(__name__)


@app.get('/')
def index():
    return render_template('index.html')


if __name__ == '__main__':
    running_on_streamlit = 'streamlit' in sys.modules or os.getenv('STREAMLIT_SERVER_PORT')
    if running_on_streamlit:
        import streamlit_app  # noqa: F401
    else:
        app.run(debug=True)
