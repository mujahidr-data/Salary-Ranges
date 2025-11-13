#!/bin/bash

# Quick push to Google Apps Script only

echo "📤 Pushing to Google Apps Script..."

# Check if clasp is installed
if ! command -v clasp &> /dev/null; then
    echo "❌ clasp not found. Installing via npx..."
    npx @google/clasp push
    exit $?
fi

# Check if .clasp.json has a valid script ID
if grep -q "YOUR_SCRIPT_ID_HERE" .clasp.json 2>/dev/null; then
    echo "❌ Please update .clasp.json with your Google Apps Script ID"
    echo ""
    echo "Option 1: Create a new Apps Script project"
    echo "   clasp create --type sheets --title 'Salary Ranges Calculator'"
    echo ""
    echo "Option 2: Use an existing project"
    echo "   Update .clasp.json with your script ID from:"
    echo "   https://script.google.com/"
    exit 1
fi

# Push
clasp push

if [ $? -eq 0 ]; then
    echo "✅ Push successful!"
else
    echo "❌ Push failed. Check clasp login status: clasp login --status"
    exit 1
fi

