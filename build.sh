#!/bin/bash
# build ทั้งไทยและอังกฤษจากไฟล์ต้นฉบับเดียว  EDMS-AutoDownload.user.js
set -e
cd /home/claude
rm -rf build-th build-en && mkdir -p build-th build-en
python3 build.py
node --check build-th/conzol.js
node --check build-th/ConZoL-Auto-Download.user.js
node --check build-en/conzol.js
node --check build-en/ConZoL-Auto-Download.user.js
echo "BUILD OK"
