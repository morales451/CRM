#!/usr/bin/env bash
# Update Roof CRM to the latest version.
# Your data lives outside this folder and is never touched.
cd "$(dirname "$0")"

if ! command -v git >/dev/null 2>&1; then
  echo "Git isn't installed. Install it, then run this again."
  exit 1
fi

if [ ! -d .git ]; then
  cat <<'MSG'
This folder was unzipped rather than cloned, so it can't update itself.
One-time fix — in the folder you want the app to live in, run:

  git clone https://github.com/morales451/CRM.git RoofCRM-App

After that, run ./update.sh inside RoofCRM-App any time.
MSG
  exit 1
fi

echo "Updating Roof CRM..."
git pull || { echo "Update failed — your data is safe, nothing changed."; exit 1; }
python3 -m pip install -q -r requirements.txt
echo "Done — start the app with ./start.sh"
