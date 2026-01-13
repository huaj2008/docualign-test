# DocuAlign - Test Version

Professional document formatting add-in for Microsoft Word.

## Features (v1.1.0)

### ✅ Working Now
- **Title Case** - Converts text with proper capitalization, preserves pharma acronyms
- **Sentence Case** - First word capitalized, preserves acronyms
- **Paragraph Support** - Works on selection OR entire paragraph when cursor is positioned

### 🔜 Coming Soon
- Title Case All Headings
- Heading Case Dialog (per-level control)
- Body Text Formatting
- Table Formatting
- Table/Figure Captions
- Cross-References
- Special Characters
- Latin Term Italicization
- TOC/LOT/LOF Insertion
- Style Scanner

## Preserved Acronyms

The following terms maintain their exact capitalization:
```
FDA, IND, NDA, BLA, ANDA, DMF, CTD, eCTD, PK, PD, ADME, AUC, Cmax, Tmax,
API, GMP, cGMP, ICH, QbD, PAT, CPP, CQA, QTPP, DNA, RNA, mRNA, pH, UV, IR,
NMR, HPLC, GC, MS, IV, IM, SC, PO, QD, BID, TID, QID, PRN, US, EU, UK,
WHO, EMA, PMDA, TGA, NMPA, SOP, QA, QC, R&D, CMC, CMO, CRO, CDMO
```

## Installation (Windows)

### Step 1: Download manifest.xml
Save the `manifest.xml` file to `C:\OfficeAddins\` (create the folder if it doesn't exist)

### Step 2: Configure Word Trust Center
1. Open Word
2. Go to **File → Options → Trust Center → Trust Center Settings**
3. Click **Trusted Add-in Catalogs**
4. In the "Catalog Url" field, enter: `\\localhost\C$\OfficeAddins\`
5. Click **Add Catalog**
6. Check ✅ **Show in Menu** for the catalog you just added
7. Click **OK** → **OK**

### Step 3: Restart Word
Close Word completely and reopen it.

### Step 4: Add the Add-in
1. Go to **Developer** tab (if not visible, enable it in Options → Customize Ribbon)
2. Click **Add-ins** → **My Add-ins**
3. Go to **Shared Folder** tab
4. Select **DocuAlign**
5. Click **Add**

## Usage

### From Ribbon
- Click the **DocuAlign** tab
- Use **Title Case** or **Sentence Case** buttons directly
- Click **More Options** to open the full panel

### From Taskpane
- Hover over any button to see detailed description
- Click to apply formatting
- Grayed buttons with "Soon" badge are coming in future updates

## Troubleshooting

### "Shared Folder" tab doesn't appear
- Make sure "Show in Menu" is checked in Trust Center
- Restart Word completely (close all Word windows)

### Add-in not loading / showing old version
Clear the Office cache:
1. Close Word
2. Delete folder: `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`
3. Restart Word

### Changes not appearing
GitHub Pages can take 1-2 minutes to update after commits.

## File Structure

```
docualign-test/
├── manifest.xml      ← Share this with users
├── taskpane.html     ← Main UI panel
├── functions.html    ← Ribbon button functions
├── assets/
│   ├── icon-16.png
│   ├── icon-32.png
│   ├── icon-64.png
│   └── icon-80.png
├── .nojekyll         ← Required for GitHub Pages
└── README.md
```

## Version History

- **v1.1.0** - Improved case conversion (paragraph support), new UI with tooltips
- **v1.0.0** - Initial release (selection-only case conversion)
