# PS-NCDU

**A PowerShell disk usage analyzer for Windows, with an interactive HTML report.**

<p align="center">
  <img src="https://img.shields.io/badge/version-3.6-2c6cb0" alt="Version">
  <img src="https://img.shields.io/badge/PowerShell-5.1%2B-5391FE?logo=powershell&logoColor=white" alt="PowerShell">
  <img src="https://img.shields.io/badge/platform-Windows-0078D6?logo=windows&logoColor=white" alt="Platform">
  <img src="https://img.shields.io/badge/license-MIT-3fa45b" alt="License">
</p>

PS-NCDU is a self-contained PowerShell script that scans a folder (or an entire drive), computes the actual size of every subfolder and file, then generates a modern, sortable, navigable **HTML report** — inspired by the Unix tool [`ncdu`](https://dev.yorhel.nl/ncdu), but designed for the Windows ecosystem and with no external dependencies.

<p align="center">
  <img src="PS-NCDU_2026-06-02.png" alt="PS-NCDU dashboard" width="100%">
</p>

<p align="center">
  <em>PowerShell&nbsp;5.1+ · Windows · Zero install · Standalone HTML report</em>
</p>

---

## Table of contents

- [Overview](#overview)
- [Features](#features)
- [Requirements](#requirements)
- [Installation](#installation)
- [Usage](#usage)
- [Parameters](#parameters)
- [The HTML report](#the-html-report)
- [Examples](#examples)
- [Troubleshooting](#troubleshooting)
- [Roadmap](#roadmap)
- [Contributing](#contributing)
- [License](#license)

---

## Overview

Pointed at a path, PS-NCDU walks the tree recursively, sums up sizes, identifies the largest folders and produces an `.html` file you can open in any browser. The report is **fully self-contained** (HTML + CSS + JavaScript in a single file): no server, no connection, nothing to install on the client side. You can archive it, email it, or drop it on a network share.

The goal: answer the question "**what is filling up this disk?**" in a few seconds, with a readable, professional look suited to enterprise use.

---

## Features

- **Recursive scan** of a folder or a drive, with configurable depth.
- **Real size computation** per folder and per file, with an aggregated total.
- **Proportion bars** color-coded by size tiers (green → amber → red) to spot the heavy hitters at a glance.
- **Sort by size** and breadcrumb navigation inside the report.
- **Light / dark theme** with a one-click toggle, using a sober, corporate palette.
- **Protected folder detection** (ACL / access denied): shown explicitly with an `ACL` badge instead of being silently skipped, with size marked `N/A`.
- **Large-file flagging** with a dedicated indicator in the stats bar.
- **File-type badges** (`.iso`, `.xlsx`, `.txt`, `.md`, …) with folder/file icons.
- **OneDrive indicator** to distinguish cloud-synced content.
- **Stats bar**: total analyzed, item count, number of large items, number of protected items.
- **100% self-contained HTML report** — a single file, openable offline.

---

## Requirements

| Item | Detail |
|---|---|
| OS | Windows 10 / 11 or Windows Server |
| PowerShell | 5.1 (Windows PowerShell) or 7+ (PowerShell Core) |
| Permissions | Read access to scanned folders; some system paths require an **administrator** console |
| Browser | Any modern browser to open the report |

No external module is required.

---

## Installation

Clone the repository or simply download the `.ps1` file:

```powershell
git clone https://github.com/Vietnamix/PS-NCDU.git
cd PS-NCDU
```

The script is encoded in **UTF-8 with BOM**: do not re-save it in another encoding, or accents and icons in the report will break.

> **Execution policy** — If Windows blocks script execution, allow it for the current session:
> ```powershell
> Set-ExecutionPolicy -Scope Process -ExecutionPolicy Bypass
> ```
> This command changes nothing permanently: it only applies to the open PowerShell window.

---

## Usage

Simplest run, on the current folder:

```powershell
.\PS-NCDU_v3_6.ps1
```

On a specific path:

```powershell
.\PS-NCDU_v3_6.ps1 -Path "C:\Users\eric"
```

The script performs the scan, generates the `.html` report and usually opens it automatically in your default browser.

---

## Parameters

> The names below describe the script's options. Adjust them if your `param()` block differs slightly.

| Parameter | Type | Description |
|---|---|---|
| `-Path` | `string` | Folder or drive to analyze. Default: the current folder. |
| `-Depth` | `int` | Maximum tree depth to traverse (e.g. `3`). |
| `-Output` | `string` | Path of the generated HTML file. Defaults to next to the script or inside the scanned folder. |
| `-MinSize` | `int` | Threshold (in MB) above which an item is flagged as "large". |
| `-Theme` | `string` | Initial report theme: `light` or `dark`. |

To display the built-in help:

```powershell
Get-Help .\PS-NCDU_v3_6.ps1 -Detailed
```

---

## The HTML report

The generated file contains:

- a **header** with the analyzed path, the date, the scan duration and the folder count;
- a **stats bar** (total, items, large, protected);
- a **sortable table**: icon, name, size, proportion bar (%), type;
- a **light/dark toggle button** in the top-right corner;
- a **footer** with the version, the scan scope and support information.

Bar colors follow size tiers to visually highlight the biggest space consumers, and inaccessible folders (ACL) stay visible with a distinct marking instead of disappearing from the report.

---

## Examples

Analyze the full user profile to 4 levels deep:

```powershell
.\PS-NCDU_v3_6.ps1 -Path "C:\Users\eric" -Depth 4
```

Analyze a whole drive and save the report to a network share:

```powershell
.\PS-NCDU_v3_6.ps1 -Path "D:\" -Output "\\server\reports\drive_D.html"
```

Start directly in dark theme:

```powershell
.\PS-NCDU_v3_6.ps1 -Path "C:\Data" -Theme dark
```

Scan system paths (administrator console recommended):

```powershell
.\PS-NCDU_v3_6.ps1 -Path "C:\Windows" -Depth 2
```

---

## Troubleshooting

| Symptom | Likely cause / fix |
|---|---|
| The script won't start | Execution policy — see the note in [Installation](#installation). |
| Broken accents or icons in the report | The `.ps1` was re-saved without a UTF-8 BOM. Restore the original encoding. |
| Many folders shown as `ACL` / `N/A` | Insufficient permissions. Re-run PowerShell **as administrator**. |
| Very long scan on a large drive | Lower `-Depth` or target a specific subfolder. |
| The report doesn't open by itself | Manually open the `.html` file shown at the end of the run. |

---

## Roadmap

- [ ] Additional CSV / JSON export of results
- [ ] Filter by file type inside the report
- [ ] Compare two scans (track changes over time)
- [ ] Live search within the table

*Suggestions welcome via issues.*

---

## Contributing

Contributions are welcome:

1. *Fork* the repository.
2. Create a branch (`git checkout -b feature/my-feature`).
3. Keep the **UTF-8 with BOM** encoding and the PowerShell *here-strings* intact.
4. Open a *pull request* clearly describing the change.

For bugs and ideas, open an **issue** stating your Windows version, PowerShell version and the command used.

---

## License

Distributed under the **MIT** license. See the [`LICENSE.md`](License.md) file.

---

<p align="center">
  <sub>PS-NCDU · Author: Eric Guiffaut · Made with PowerShell 💙</sub>
</p>
