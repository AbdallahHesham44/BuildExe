#!/usr/bin/env python3
"""
Automated Update and Build Script for Cross Reference Converter
---------------------------------------------------------------
Usage:
    python update_and_build.py ["Your commit message"]
    or:
    ./update_and_build.sh ["Your commit message"]

What this script does:
1. Syncs any edits in /kaggle/working/cross_converter.py into the repo.
2. Runs the unit tests (test_converter.py).
3. Commits and pushes changes to GitHub (AbdallahHesham44/BuildExe).
4. Monitors GitHub Actions until the Windows .exe finishes compiling.
5. Automatically downloads the compiled CrossConverter.exe / zip to /kaggle/working/!
"""

import os
import sys
import time
import json
import shutil
import zipfile
import subprocess
import urllib.request
import urllib.error
from datetime import datetime

# ----------------- CONFIGURATION -----------------
GITHUB_REPO = "AbdallahHesham44/BuildExe"
GITHUB_BRANCH = "main"
WORKFLOW_NAME = "build_cross_converter.yml"
ARTIFACT_NAME = "CrossConverter-Windows-exe"

PROJECT_ROOT = "/kaggle/working/projects"
SRC_CONVERTER = "/kaggle/working/cross_converter.py"
REPO_CONVERTER = "/kaggle/working/projects/cross_converter/cross_converter.py"
OUTPUT_DIR = "/kaggle/working"
TOKEN_PATHS = [
    "/kaggle/working/.github_token",
    os.path.join(PROJECT_ROOT, ".github_token"),
    os.path.expanduser("~/.github_token"),
]
# -------------------------------------------------


def get_token():
    token = os.environ.get("GITHUB_TOKEN")
    if token:
        return token.strip()
    for p in TOKEN_PATHS:
        if os.path.exists(p):
            try:
                with open(p, "r") as f:
                    t = f.read().strip()
                    if t:
                        return t
            except Exception:
                pass
    err("GitHub token not found! Please place it in /kaggle/working/.github_token or set GITHUB_TOKEN env var.")


def log(step, msg):
    print(f"\n\033[1;34m[{step}]\033[0m \033[1m{msg}\033[0m")


def info(msg):
    print(f"    \033[0;32m✓\033[0m {msg}")


def warn(msg):
    print(f"    \033[0;33m⚠\033[0m {msg}")


def err(msg):
    print(f"\n\033[1;31m[ERROR]\033[0m {msg}")
    sys.exit(1)


def run_cmd(cmd, cwd=PROJECT_ROOT, check=True):
    res = subprocess.run(cmd, cwd=cwd, shell=isinstance(cmd, str), capture_output=True, text=True)
    if check and res.returncode != 0:
        err(f"Command failed: {cmd}\nOutput:\n{res.stdout}\n{res.stderr}")
    return res


def api_request(url, method="GET", data=None, token=None):
    if not token:
        token = get_token()
    headers = {
        "Authorization": f"token {token}",
        "Accept": "application/vnd.github.v3+json",
        "User-Agent": "BuildExe-Updater",
    }
    body = json.dumps(data).encode("utf-8") if data else None
    req = urllib.request.Request(url, data=body, headers=headers, method=method)
    try:
        with urllib.request.urlopen(req, timeout=30) as resp:
            content = resp.read()
            return json.loads(content.decode("utf-8")) if content else {}
    except urllib.error.HTTPError as e:
        err(f"GitHub API Error [{e.code}] for {url}: {e.read().decode('utf-8', errors='ignore')}")
    except Exception as e:
        err(f"Request failed for {url}: {e}")


def sync_files():
    log("1/5", "Checking and syncing project files...")

    # If /kaggle/working/cross_converter.py is edited, sync it into the repo
    if os.path.exists(SRC_CONVERTER) and os.path.exists(REPO_CONVERTER):
        with open(SRC_CONVERTER, "rb") as f1, open(REPO_CONVERTER, "rb") as f2:
            if f1.read() != f2.read():
                shutil.copy2(SRC_CONVERTER, REPO_CONVERTER)
                info(f"Synced changes from {SRC_CONVERTER} -> {REPO_CONVERTER}")
            else:
                info("Source files are already in sync.")
    elif os.path.exists(SRC_CONVERTER) and not os.path.exists(REPO_CONVERTER):
        os.makedirs(os.path.dirname(REPO_CONVERTER), exist_ok=True)
        shutil.copy2(SRC_CONVERTER, REPO_CONVERTER)
        info(f"Copied {SRC_CONVERTER} to repository.")


def run_tests():
    log("2/5", "Running unit tests...")
    test_path = os.path.join(PROJECT_ROOT, "cross_converter", "test_converter.py")
    test_dir = os.path.dirname(test_path)

    python_bin = sys.executable
    venv_python = "/kaggle/working/.venv/bin/python"
    if os.path.exists(venv_python):
        python_bin = venv_python

    res = subprocess.run([python_bin, "test_converter.py"], cwd=test_dir, capture_output=True, text=True)
    if res.returncode != 0:
        err(f"Unit tests failed! Please fix errors before pushing:\n{res.stderr or res.stdout}")
    info("All unit tests passed successfully!")


def commit_and_push(commit_msg):
    log("3/5", "Committing and pushing to GitHub...")
    token = get_token()

    run_cmd(["git", "add", "-A"], cwd=PROJECT_ROOT)
    diff = run_cmd(["git", "status", "--porcelain"], cwd=PROJECT_ROOT).stdout.strip()

    if not diff:
        warn("No new code changes to commit.")
        commit_sha = run_cmd(["git", "rev-parse", "HEAD"], cwd=PROJECT_ROOT).stdout.strip()
        info(f"Using current HEAD commit: {commit_sha[:7]}")
        return commit_sha, False

    run_cmd(["git", "config", "user.name", "Abdallah Hesham"], cwd=PROJECT_ROOT)
    run_cmd(["git", "config", "user.email", "124697660+AbdallahHesham44@users.noreply.github.com"], cwd=PROJECT_ROOT)

    run_cmd(["git", "commit", "-m", commit_msg], cwd=PROJECT_ROOT)
    commit_sha = run_cmd(["git", "rev-parse", "HEAD"], cwd=PROJECT_ROOT).stdout.strip()
    info(f"Created commit: [{commit_sha[:7]}] {commit_msg}")

    push_url = f"https://x-access-token:{token}@github.com/{GITHUB_REPO}.git"
    run_cmd(["git", "push", push_url, f"HEAD:{GITHUB_BRANCH}"], cwd=PROJECT_ROOT)
    info(f"Pushed to https://github.com/{GITHUB_REPO} ({GITHUB_BRANCH})")
    return commit_sha, True


def wait_for_build(commit_sha):
    log("4/5", "Monitoring GitHub Actions Windows build...")
    token = get_token()

    run_id = None
    start_time = time.time()
    info(f"Waiting for GitHub Actions build for commit [{commit_sha[:7]}]...")

    # 1. First poll for an auto-triggered push workflow matching this commit
    for _ in range(8):  # wait up to 40s
        time.sleep(5)
        runs_data = api_request(f"https://api.github.com/repos/{GITHUB_REPO}/actions/runs?per_page=5", token=token)
        for r in runs_data.get("workflow_runs", []):
            if r.get("head_sha") == commit_sha:
                run_id = r["id"]
                info(f"Build workflow detected! Run ID: {run_id}")
                info(f"Tracking run URL: {r.get('html_url')}")
                break
        if run_id:
            break

    # 2. If no push workflow triggered (e.g. non-source file changed), dispatch manually
    if not run_id:
        info("No automatic push run detected; triggering workflow manually via workflow_dispatch...")
        api_request(
            f"https://api.github.com/repos/{GITHUB_REPO}/actions/workflows/{WORKFLOW_NAME}/dispatches",
            method="POST",
            data={"ref": GITHUB_BRANCH},
            token=token,
        )
        for _ in range(8):
            time.sleep(5)
            runs_data = api_request(f"https://api.github.com/repos/{GITHUB_REPO}/actions/runs?per_page=3", token=token)
            for r in runs_data.get("workflow_runs", []):
                if r.get("status") in ["queued", "in_progress"]:
                    run_id = r["id"]
                    info(f"Dispatched build workflow found! Run ID: {run_id}")
                    info(f"Tracking run URL: {r.get('html_url')}")
                    break
            if run_id:
                break

    if not run_id:
        err("Could not find or trigger a workflow run on GitHub Actions. Please check your GitHub repository.")

    # 3. Wait for completion
    last_status = ""
    while True:
        data = api_request(f"https://api.github.com/repos/{GITHUB_REPO}/actions/runs/{run_id}", token=token)
        status = data.get("status")
        conclusion = data.get("conclusion")

        status_display = f"{status}" + (f" -> {conclusion}" if conclusion else "")
        if status_display != last_status:
            info(f"Build status: {status_display} (elapsed: {int(time.time() - start_time)}s)")
            last_status = status_display

        if status == "completed":
            if conclusion == "success":
                info("Windows executable build completed successfully!")
                return run_id
            else:
                err(f"Build finished with failure conclusion: {conclusion}.\nCheck logs at: {data.get('html_url')}")

        time.sleep(10)


def download_artifact(run_id):
    log("5/5", "Downloading compiled executable artifact...")
    token = get_token()

    artifacts_url = f"https://api.github.com/repos/{GITHUB_REPO}/actions/runs/{run_id}/artifacts"
    data = api_request(artifacts_url, token=token)
    artifacts = data.get("artifacts", [])

    matched = None
    for a in artifacts:
        if a["name"] == ARTIFACT_NAME:
            matched = a
            break

    if not matched and artifacts:
        matched = artifacts[0]

    if not matched:
        err("No artifacts found for this workflow run. Check the GitHub Actions run logs.")

    download_url = matched["archive_download_url"]
    dest_zip = os.path.join(OUTPUT_DIR, f"{ARTIFACT_NAME}.zip")
    info(f"Downloading {matched['name']} ({matched['size_in_bytes'] / (1024*1024):.1f} MB)...")

    cmd = [
        "curl", "-s", "-L",
        "-H", f"Authorization: token {token}",
        download_url,
        "-o", dest_zip,
    ]
    res = subprocess.run(cmd, capture_output=True)
    if res.returncode != 0 or not os.path.exists(dest_zip):
        err(f"Failed to download artifact via curl: {res.stderr}")

    info(f"Saved artifact zip to: {dest_zip}")

    # Extract CrossConverter.exe into OUTPUT_DIR
    try:
        with zipfile.ZipFile(dest_zip, "r") as z:
            z.extractall(OUTPUT_DIR)
        exe_path = os.path.join(OUTPUT_DIR, "CrossConverter.exe")
        if os.path.exists(exe_path):
            info(f"Extracted executable to: {exe_path}")
    except Exception as e:
        warn(f"Could not extract zip: {e}")

    return dest_zip


def main():
    commit_msg = sys.argv[1] if len(sys.argv) > 1 else f"Update cross_converter - {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}"

    print("\033[1;36m========================================================\033[0m")
    print("\033[1;36m   Cross Reference Converter - Auto Update & Build      \033[0m")
    print("\033[1;36m========================================================\033[0m")

    sync_files()
    run_tests()
    commit_sha, has_changes = commit_and_push(commit_msg)
    run_id = wait_for_build(commit_sha)
    zip_path = download_artifact(run_id)

    print("\n\033[1;32m========================================================\033[0m")
    print("\033[1;32m   🎉 SUCCESS! Windows Executable is Ready to Use!     \033[0m")
    print("\033[1;32m========================================================\033[0m")
    print(f"• Download Zip:  \033[1m{zip_path}\033[0m")
    print(f"• Direct Exe:    \033[1m{os.path.join(OUTPUT_DIR, 'CrossConverter.exe')}\033[0m")
    print(f"• GitHub Run:    https://github.com/{GITHUB_REPO}/actions/runs/{run_id}")
    print("\nYou can now download CrossConverter.exe directly and run it on Windows!\n")


if __name__ == "__main__":
    main()
