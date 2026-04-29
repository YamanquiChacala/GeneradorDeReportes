#!/bin/bash

# Exit immediately if a command fails
set -e

echo "🚀 Initializing new Google Apps Script project..."

# ---------------------------------------------------------
# 1. GIT LFS SETUP
# ---------------------------------------------------------
echo "📦 Configuring Git LFS..."
if command -v git-lfs >/dev/null 2>&1 || git lfs --version >/dev/null 2>&1; then
  git lfs install
  echo "✅ Git LFS initialized."
else
  echo "⚠️ Git LFS is not installed on your system. Please install it (e.g., sudo apt install git-lfs) to use LFS features."
fi

# ---------------------------------------------------------
# 2. CLASP SETUP
# ---------------------------------------------------------
echo "🔗 Let's link this to Google Apps Script."
echo "Do you want to create a [N]ew standalone project or [L]ink to an existing one? (N/l)"
read -r CLASP_CHOICE

if [[ "$CLASP_CHOICE" =~ ^[Ll]$ ]]; then
  echo "Please enter the existing Script ID (found in Project Settings in the Apps Script editor):"
  read -r SCRIPT_ID
  
  # Manually create the .clasp.json to avoid 'clasp clone' overwriting your local boilerplate files
  echo '{"scriptId":"'"$SCRIPT_ID"'","rootDir":"./dist"}' > .clasp.json
  echo "✅ Linked to existing project."

else
  echo "Creating a new standalone Apps Script project..."
  # This will prompt you to choose the script type (standalone, docs, sheets, etc.)
  npx clasp create
  
  # Clasp creates .clasp.json without our required rootDir. 
  # This quick Node snippet safely injects "rootDir": "./dist" into the JSON file.
  node -e "
    const fs = require('fs');
    const config = JSON.parse(fs.readFileSync('.clasp.json'));
    config.rootDir = './dist';
    fs.writeFileSync('.clasp.json', JSON.stringify(config, null, 2));
  "
  
  # Clasp create also generates a default appsscript.json in the root. 
  # We delete it because your true manifest lives in src/appsscript.json
  rm -f appsscript.json
  
  echo "✅ Created new project and configured rootDir to ./dist."
fi

# ---------------------------------------------------------
# 3. INITIAL BUILD
# ---------------------------------------------------------
echo "🛠️  Running initial build..."
npm run build

echo "🎉 All set! Your project is ready. Use 'npm run push' to deploy your code."