#!/usr/bin/env bash
# setup-repo.sh — run this AFTER creating the empty repo on GitHub:
#   https://github.com/new  ->  name: ap-automation-portfolio, public, no README
set -euo pipefail

REPO="https://github.com/andry-000/ap-automation-portfolio.git"

if [ ! -f "Code.gs" ] || [ ! -f "appsscript.json" ]; then
  echo "Run this from the project root (where Code.gs lives)."
  exit 1
fi

if [ ! -d ".git" ]; then
  git init -b main
  git config user.email  "${GIT_EMAIL:-you@example.com}"
  git config user.name   "${GIT_NAME:-Your Name}"
fi

git add -A
if ! git diff --cached --quiet; then
  git commit -m "feat: initial AP invoices dashboard with resumable Gmail ingestion"
fi

if ! git remote get-url origin >/dev/null 2>&1; then
  git remote add origin "$REPO"
fi

git branch -M main
git push -u origin main
echo "Done. Visit $REPO"
