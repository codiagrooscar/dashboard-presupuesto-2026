import os
import re
import subprocess

node_exe = r'C:\Users\oscar.ocampo\.cache\codex-runtimes\codex-primary-runtime\dependencies\node\bin\node.exe'

def check_html(path):
    print(f"Checking HTML file: {path}")
    with open(path, 'r', encoding='utf-8') as f:
        content = f.read()
    pattern = re.compile(r'<script(?:\s+type=[\"\'](?:text/javascript|module)[\"\'])?>(.*?)</script>', re.DOTALL | re.IGNORECASE)
    scripts = pattern.findall(content)
    print(f"  Found {len(scripts)} inline script blocks.")
    for idx, s in enumerate(scripts):
        if not s.strip():
            continue
        tmp = f"tmp_check_{idx}.js"
        with open(tmp, 'w', encoding='utf-8') as tf:
            tf.write(s)
        res = subprocess.run([node_exe, '--check', tmp], capture_output=True, text=True)
        if os.path.exists(tmp):
            os.remove(tmp)
        if res.returncode != 0:
            print(f"  FAIL in block {idx}:")
            print(res.stderr[:400])
            return False
        else:
            print(f"  Block {idx}: OK")
    return True

def check_folder_js(folder):
    print(f"Checking folder .js files: {folder}")
    for root, dirs, files in os.walk(folder):
        for f in files:
            if f.endswith('.js'):
                p = os.path.join(root, f)
                res = subprocess.run([node_exe, '--check', p], capture_output=True, text=True)
                if res.returncode != 0:
                    print(f"  FAIL in {p}: {res.stderr[:300]}")
                    return False
                else:
                    print(f"  {f}: OK")
    return True

if __name__ == '__main__':
    ok1 = check_html('web_dashboard/index.html')
    ok2 = check_html('web_dashboard_direccion/index.html')
    ok3 = check_folder_js('web_dashboard_direccion')
    if ok1 and ok2 and ok3:
        print("\n>>> ALL CHECKS PASSED: 100% VALID JAVASCRIPT EVERYWHERE! <<<")
    else:
        exit(1)
