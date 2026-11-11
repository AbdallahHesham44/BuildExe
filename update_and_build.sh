#!/usr/bin/env bash
set -e

# Run the automated update and build script
python3 /kaggle/working/update_and_build.py "$@"
