#!/bin/bash
cd /opt/asistan || exit 1
git add -A
git diff --cached --quiet || git commit -qm "otomatik yedek $(date '+%Y-%m-%d %H:%M')"
git push -q origin main
