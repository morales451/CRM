#!/usr/bin/env bash
# Update Roof CRM now. ./start.sh already does this on every start.
# Your data lives outside this folder and is never touched.
{
  cd "$(dirname "$0")"
  python3 updater.py
  exit
}
