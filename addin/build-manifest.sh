#!/bin/sh
# Usage: ./build-manifest.sh https://<user>.github.io/<repo>/addin
set -e
BASE="${1%/}"
ORIGIN=$(printf '%s' "$BASE" | sed -E 's#^(https://[^/]+).*#\1#')
sed -e "s#{{BASE}}#$BASE#g" -e "s#{{ORIGIN}}#$ORIGIN#g" "$(dirname "$0")/manifest.template.xml" > "$(dirname "$0")/manifest.xml"
echo "manifest.xml -> $BASE"
