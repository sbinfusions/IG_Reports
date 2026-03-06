# GitHub Pages Setup Guide

## ✅ What's Been Updated

Your `run_report.bat` script is now configured to push reports to:
- **Repository**: `https://github.com/sbinfusions/IG_Reports`
- **Branch**: `gh-pages`
- **Live URL**: `https://sbinfusions.github.io/IG_Reports/`

## 🚀 One-Time Setup Steps

### Step 1: Create the gh-pages Branch

Run these commands in the `templates` folder:

```bash
cd templates
git checkout --orphan gh-pages
git rm -rf .
git commit --allow-empty -m "Initialize GitHub Pages"
git push -u origin gh-pages
git checkout master
```

**OR use this simpler method:**

```bash
cd templates
git checkout -b gh-pages
git push -u origin gh-pages
git checkout master
```

### Step 2: Enable GitHub Pages (if needed)

1. Go to: https://github.com/sbinfusions/IG_Reports
2. Click **Settings** → **Pages**
3. Verify it says: "Your site is published at https://sbinfusions.github.io/IG_Reports/"
4. GitHub should auto-detect the `gh-pages` branch

## 📝 How to Use

1. Run `run_report.bat`
2. Select your Word document
3. Report generates and opens in browser
4. When prompted, type `y` to push to GitHub Pages
5. Your report will be live at:
   ```
   https://sbinfusions.github.io/IG_Reports/[filename]_report.html
   ```

## 🔄 Monthly Workflow

Each month:
1. Generate new report with `run_report.bat`
2. Push to GitHub Pages (all previous reports remain accessible)
3. Share the URL with your team/client

## 📂 File Structure on gh-pages

```
gh-pages branch:
├── CONTENT_SCHEDULE_JAN_report.html
├── CONTENT_SCHEDULE_FEB_report.html
├── CONTENT_SCHEDULE_MAR_report.html
├── images/
│   └── [all downloaded images]
└── *.png (brand assets)
```

## ⚠️ Important Notes

- The `gh-pages` branch is **separate** from your main code
- Old reports are **never deleted** (great for archives!)
- Images are stored in the `images/` subfolder automatically
- Each push adds/updates files without removing old ones

## 🐛 Troubleshooting

**"Push failed" error?**
- Make sure you've created the `gh-pages` branch (see Step 1)
- Check your GitHub authentication: `git config --global user.name`

**Report not appearing on GitHub Pages?**
- Wait 1-2 minutes for GitHub to publish
- Check Settings → Pages is enabled
- Verify the branch is `gh-pages`

**Images not loading?**
- The script downloads images to `Content_Schedule_OUTPUT/images/`
- They're automatically included in the git push
- If using Dropbox links, they may expire (use downloaded images instead)
