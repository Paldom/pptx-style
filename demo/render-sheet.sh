#!/usr/bin/env bash
# Regenerates demo/ops-review-contact-sheet.png: the 18-slide deck from the demo,
# rendered with LibreOffice using the bundled OFL test font. Run from the repo root.
set -euo pipefail
source demo/setup.sh
python3 "$PROFILE/profile.py" messy.potx >/dev/null
python3 "$BUILD/build.py" spec.json >/dev/null
python3 "$BUILD/render.py" ops-review.pptx --contract .pptx-style/messy/contract.json --require
cp ops-review-render/contact-sheet.png "$REPO/demo/ops-review-contact-sheet.png"
