# Roadmap / ideas

Comment on (or open) an issue first so we don't duplicate work.

## Good first issues
- [x] Add screenshots (with sample coins) to the README
- [x] Ship a small sample collection (`samples/SampleCollection.csv`) so new users can explore
- [ ] Convert the plain-text manual to Markdown
- [ ] Add a `-DataPath` parameter so the collection can live outside the script folder

## Features
- [ ] Live spot-price lookup for the melt-value calculator
- [ ] Import from other collection tools (Numista, CoinManage) via CSV
- [ ] World coins, currency and bullion templates
- [ ] Dark theme
- [ ] Export the collection to JSON for use in other apps

## Quality
- [ ] Split the 4,000-line script into a module (data / UI / reports)
- [ ] Pester tests for the CSV import, melt-value maths and duplicate detection
