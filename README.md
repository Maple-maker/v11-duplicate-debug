# DD1750 Converter - Fixed Version

This application converts BOM (Bill of Materials) PDFs into DD1750 forms with all bugs fixed:

## ✅ Fixed Issues:
1. **18 items per page** (was 40 - wrong)
2. **Correct NSN assignment** (Item 1 gets its own NSN, not Item 2's)
3. **Clean descriptions** (removes material IDs like C_75Q65)

## How to Deploy on Railway:

### Step 1: Create GitHub Repository
1. Go to https://github.com
2. Click "New repository"
3. Name it "dd1750-fixed"
4. Click "Create repository"

### Step 2: Upload Files to GitHub
1. In your new repository, click "Add file" → "Upload files"
2. Drag and drop all 5 files from your folder:
   - `app.py`
   - `dd1750_core.py`
   - `requirements.txt`
   - `Procfile`
   - `README.md`
3. Click "Commit changes"

### Step 3: Deploy on Railway
1. Go to https://railway.app
2. Click "New Project"
3. Select "Deploy from GitHub repo"
4. Find and select your "dd1750-fixed" repository
5. Railway will automatically deploy your app!

## How to Use:
1. Visit your Railway app URL
2. Upload your BOM PDF (like B49.pdf)
3. Upload your blank DD1750 template PDF
4. Click "Generate DD1750"
5. Download the filled form

## Verify Fixes Work:
- Open the generated PDF
- Count items on page 1: Should be **18 maximum**
- Check Item 1: Should have its **own NSN**
- Check descriptions: Should be **clean** (no C_75Q65 codes)
