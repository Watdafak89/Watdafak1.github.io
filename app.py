import os
import runpy
import sys
from pathlib import Path

from flask import Flask, render_template

app = Flask(__name__)


@app.get('/')
def index():
    return render_template('index.html')


if __name__ == '__main__':
    running_on_streamlit = 'streamlit' in sys.modules or os.getenv('STREAMLIT_SERVER_PORT')
    if running_on_streamlit:
        # Streamlit reruns this entrypoint after widget changes. A normal import
        # is cached, leaving the page blank on uploads and other interactions.
        runpy.run_path(str(Path(__file__).with_name('streamlit_app.py')), run_name='__main__')
    else:
        app.run(debug=True)
