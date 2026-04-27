import os
import subprocess
import sys
import tempfile
import time

python_exe = sys.executable
with tempfile.NamedTemporaryFile('w', suffix='.py', delete=False, dir=os.getcwd(), encoding='utf-8') as f:
    f.write('import streamlit as st\n')
    f.write('st.title(\"Test Headless\")\n')
    app_path = f.name

print('APP', app_path)
env = os.environ.copy()
env['STREAMLIT_SERVER_HEADLESS'] = 'true'
env['STREAMLIT_SERVER_RUN_ON_SAVE'] = 'false'
env['STREAMLIT_BROWSER_GATHER_USAGE_STATS'] = 'false'
env['STREAMLIT_SERVER_FILE_WATCHER_TYPE'] = 'none'

proc = subprocess.Popen(
    [python_exe, '-m', 'streamlit', 'run', app_path, '--server.port', '8509', '--server.headless', 'true', '--server.runOnSave', 'false'],
    stdout=subprocess.PIPE,
    stderr=subprocess.PIPE,
    env=env,
    text=True,
)

time.sleep(6)
print('POLL', proc.poll())
out, err = proc.communicate(timeout=10)
print('OUT----\n', out)
print('ERR----\n', err)
proc.terminate()
proc.wait(timeout=5)
os.unlink(app_path)
