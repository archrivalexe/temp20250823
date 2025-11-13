# Experimental Design & Procedure Figure (APA-friendly, PDF 1200px width)

This repo contains:
- Mermaid source for the figure (APA-friendly styling).
- A script to export a 1200 px wide PDF using mermaid-cli.
- An APA-style caption file.
- An optional GitHub Actions workflow to auto-build the PDF on push/PR.

## Quick start (local export)
1) Install Node.js (>= 18) and run:
   npm install
2) Export PDF:
   npm run build:pdf

Output will be in build/figure-procedure-flowchart.pdf (vector PDF, width target = 1200 px).

## Notes for APA (7th)
- Monochrome linework, high-contrast labels, sans-serif figure font per APA guidance for figures.
- Caption provided in figure-caption-APA.md (number, italicized title, and Note).
- If your journal requires single- or double-column fit, place/scale the PDF in your manuscript accordingly.

## GitHub Actions (optional)
If you commit to a GitHub repo with the provided workflow, every push/PR will generate build/figure-procedure-flowchart.pdf automatically and upload it as an artifact.

## Command details
We use mermaid-cli (mmdc):
- Width is fixed at 1200 px for export (--width 1200).
- PDF is vector; it scales cleanly in layout tools.
- Theme is neutral; font family set to Arial/Helvetica for figure clarity.