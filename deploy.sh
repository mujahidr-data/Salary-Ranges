#!/bin/bash

# Salary Ranges - Deploy Script
# Pushes to Google Apps Script and commits to Git

echo "🚀 Deploying Salary Ranges Calculator..."

# Check if clasp is installed
if ! command -v clasp &> /dev/null; then
    echo "❌ clasp is not installed. Install it with: npm install -g @google/clasp"
    exit 1
fi

# Check if logged in to clasp
if ! clasp login --status &> /dev/null; then
    echo "❌ Not logged in to clasp. Run: clasp login"
    exit 1
fi

# Check if .clasp.json has a valid script ID
if grep -q "YOUR_SCRIPT_ID_HERE" .clasp.json 2>/dev/null; then
    echo "❌ Please update .clasp.json with your Google Apps Script ID"
    echo "   Run: clasp create --type sheets --title 'Salary Ranges Calculator'"
    echo "   Or update .clasp.json manually with your existing script ID"
    exit 1
fi

# Push to Apps Script
echo "📤 Pushing to Google Apps Script..."
if clasp push; then
    echo "✅ Apps Script push successful"
else
    echo "❌ Apps Script push failed"
    exit 1
fi

# Commit to Git
echo "📝 Committing to Git..."
git add .
git commit -m "Deploy: $(date '+%Y-%m-%d %H:%M:%S')"

# Push to Git
echo "🔼 Pushing to Git remote..."
if git push; then
    echo "✅ Git push successful"
else
    echo "⚠️  Git push failed (this is okay if you haven't set up a remote yet)"
fi

echo "✨ Deployment complete!"

