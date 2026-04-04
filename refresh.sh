#!/usr/bin/env bash
# Rebuild the KPI dashboard from the OneDrive CSV and push to GitHub.
# Run this from the KPIdashboard folder whenever you want to publish fresh data.
#
#   ./refresh.sh
#
set -e

REPO_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$REPO_DIR"

echo "Building dashboard from OneDrive CSV..."
node build.js --output public/index.html

echo "Committing and pushing..."
git add public/index.html
if git diff --staged --quiet; then
  echo "No changes — dashboard is already up to date."
else
  git commit -m "chore: refresh dashboard $(date '+%Y-%m-%d')"
  git push
  echo "Done. Dashboard updated at https://francescop85-max.github.io/kpi-dashboard/"
fi
