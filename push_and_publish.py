# -*- coding: utf-8 -*-
import sys
import io
import os
import subprocess
import urllib.request
import json

sys.stdout = io.TextIOWrapper(sys.stdout.buffer, encoding='utf-8')

token = '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'
owner = 'Radwakamalhamdy'
repo = 'Radwa-Masters-PHD-'

headers = {
    'Authorization': f'Bearer {token}',
    'Accept': 'application/vnd.github+json',
    'User-Agent': 'Antigravity-Agent'
}

print("1. Initializing Git repository locally...")
subprocess.run(['git', 'init'], check=True)
subprocess.run(['git', 'config', 'user.name', 'Radwa Kamal'], check=True)
subprocess.run(['git', 'config', 'user.email', 'radwakamal@users.noreply.github.com'], check=True)

# Rename branch to main
subprocess.run(['git', 'branch', '-M', 'main'], check=True)

# Add files
subprocess.run(['git', 'add', '-A'], check=True)

# Commit
commit_res = subprocess.run(['git', 'commit', '-m', 'Initial release: 148 Master and PhD Scholarships and Interactive Hub'], capture_output=True, text=True)
print("Commit output:", commit_res.stdout)

# Setup remote with authentication
remote_url = f"https://x-access-token:{token}@github.com/{owner}/{repo}.git"

subprocess.run(['git', 'remote', 'remove', 'origin'], stderr=subprocess.DEVNULL)
subprocess.run(['git', 'remote', 'add', 'origin', remote_url], check=True)

print("2. Pushing to GitHub main branch...")
push_res = subprocess.run(['git', 'push', '-u', 'origin', 'main', '--force'], capture_output=True, text=True)
print("Push output:", push_res.stdout, push_res.stderr)

# 3. Make repo public so anyone can access it as requested
print("3. Updating repository visibility to public...")
patch_data = json.dumps({"private": False}).encode('utf-8')
req = urllib.request.Request(f"https://api.github.com/repos/{owner}/{repo}", data=patch_data, headers=headers, method='PATCH')
try:
    with urllib.request.urlopen(req) as resp:
        res = json.loads(resp.read().decode('utf-8'))
        print(f"Repo visibility updated: private={res.get('private')}, html_url={res.get('html_url')}")
except Exception as e:
    print(f"Could not change visibility: {e}")

# 4. Enable GitHub Pages
print("4. Enabling GitHub Pages on main branch...")
pages_data = json.dumps({
    "source": {
        "branch": "main",
        "path": "/"
    }
}).encode('utf-8')

req_pages = urllib.request.Request(f"https://api.github.com/repos/{owner}/{repo}/pages", data=pages_data, headers=headers, method='POST')
try:
    with urllib.request.urlopen(req_pages) as resp:
        res_pages = json.loads(resp.read().decode('utf-8'))
        print(f"GitHub Pages enabled! URL: {res_pages.get('html_url')}")
except urllib.error.HTTPError as e:
    if e.code == 409 or e.code == 400: # already enabled or conflict
        print("Pages might already be enabled or being built...")
    else:
        print(f"Pages status: {e}")
except Exception as e:
    print(f"Pages note: {e}")

# 5. Sanitize local git remote so token is not saved in plain text in config
clean_remote = f"https://github.com/{owner}/{repo}.git"
subprocess.run(['git', 'remote', 'set-url', 'origin', clean_remote], check=True)
print("Sanitized remote URL to clean HTTPS.")

# Remove this script after execution
if os.path.exists(__file__):
    os.remove(__file__)
    print("Cleaned up script.")
