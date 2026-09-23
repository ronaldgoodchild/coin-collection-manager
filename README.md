# Coin Collection Manager

A free **coin collection tracker** for Windows: a full mouse-driven WPF dashboard written as a single PowerShell script. Track coins, slabs and certifications, attach photos, print flip labels, see what your collection is worth, and produce an insurance-ready report. No install, no database, no cloud - your data stays in plain CSV files on your PC.

> Built by a working IT technician for a family coin collection. Free to use, free to change.

## Features

- **Dashboard** of stat cards and a sortable, filterable coin grid; multiple collections with a switcher
- **Add / edit coins** - denomination, year, mint mark, grade, variety/error, notes, slab and certification tracking
- **Coin images** - attach photos to any coin
- **PCGS / NGC certification lookup** and one-click **online research links** (price guides, population reports, auction results)
- **Melt-value calculator** for silver and gold coins
- **Reports** (HTML / text, printable): insurance / appraisal report, missing coins, duplicates, holdings, set-completion tracker, value history and charts
- **2x2 flip-label printing**
- **Backup & restore**, CSV import/export, and **Excel import** from a spreadsheet-based collection
- **Duplicate detection**, variety and error coin tracking, value-history tracking
- Keyboard shortcuts and right-click context menus

See [docs/USER_MANUAL.txt](docs/USER_MANUAL.txt) for the full 28-section manual.

## Requirements

- Windows 10 / 11 with Windows PowerShell 5.1 (built in) and .NET Framework 4.5+
- Microsoft Excel - only if you want to import an existing Excel collection

## Quick start

```powershell
git clone https://github.com/ronaldgoodchild/coin-collection-manager.git
cd coin-collection-manager
powershell -ExecutionPolicy Bypass -File .\CoinCollection.ps1
```

On first launch it creates `Collections\MyCoins.csv` and compiles a small image helper (`ImagePathConverter.dll`) next to the script.

## Your data

Everything is stored next to the script: `Collections\` (CSV), `CoinImages\`, `Backups\`, `Reports\`, `ValueHistory.csv`. These folders are git-ignored - **never commit them** (your inventory and appraisal reports are sensitive). Use the built-in Backup to keep copies.

## Contributing

Ideas and pull requests welcome - see [CONTRIBUTING.md](CONTRIBUTING.md) and [ROADMAP.md](ROADMAP.md).

## License

[MIT](LICENSE) (c) 2026 Ronald Goodchild / REGTeches
