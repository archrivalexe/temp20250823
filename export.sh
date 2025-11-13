#!/usr/bin/env bash
set -euo pipefail
mkdir -p build
npx -y @mermaid-js/mermaid-cli@10.9.1 \
  -i figure-procedure-flowchart.mmd \
  -o build/figure-procedure-flowchart.pdf \
  -t neutral \
  -w 1200 \
  --pdfFit \
  -b transparent
echo "Done: build/figure-procedure-flowchart.pdf"