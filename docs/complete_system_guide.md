# Voter Analytics - Complete System Guide

## Overview

A cloud-based system to extract voter data from Tamil Nadu Electoral Roll PDFs using OCR.

```
┌──────────────────────────────────────────────────────────────────┐
│                        SYSTEM ARCHITECTURE                        │
├──────────────────────────────────────────────────────────────────┤
│                                                                   │
│   YOUR PC                        CLOUD VM (GCP)                  │
│   ────────                       ──────────────                  │
│                                                                   │
│   ┌─────────┐    Upload PDFs    ┌─────────────┐                 │
│   │ Browser │ ───────────────→  │ Streamlit   │                 │
│   │         │                   │ Web UI      │                 │
│   └─────────┘                   └──────┬──────┘                 │
│        ▲                               │                         │
│        │                               ▼                         │
│        │                        ┌─────────────┐                 │
│        │                        │ PDF → Image │                 │
│        │                        │ Extraction  │                 │
│        │                        └──────┬──────┘                 │
│        │                               │                         │
│        │                               ▼                         │
│        │                        ┌─────────────┐                 │
│        │                        │ Tesseract   │                 │
│        │                        │ OCR Engine  │                 │
│        │                        └──────┬──────┘                 │
│        │                               │                         │
│        │                               ▼                         │
│        │                        ┌─────────────┐                 │
│        │   Download Excel       │ Excel       │                 │
│        │ ◄───────────────────── │ Generator   │                 │
│        │                        └──────┬──────┘                 │
│        │                               │                         │
│   ┌────┴────┐                          ▼                         │
│   │ 📱 Ntfy │ ◄──────────────── 🔔 Notification                 │
│   │ App     │   Push Alert                                       │
│   └─────────┘                                                    │
│                                                                   │
└──────────────────────────────────────────────────────────────────┘
```

---

## Step-by-Step Process

### Phase 1: One-Time Setup (10 minutes)

#### 1.1 Create GCP Account

1. Go to [console.cloud.google.com](https://console.cloud.google.com)
2. Sign in with Google account
3. Activate **Free $300 credit** (valid for 90 days)

#### 1.2 Create VM Instance

1. Go to **Compute Engine → VM Instances**
2. Click **Create Instance**
3. Configure:

| Setting | Value |
|---------|-------|
| Name | `voter-analytics-vm` |
| Region | `asia-south1 (Mumbai)` |
| Machine Type | `e2-standard-8` (8 vCPU, 32GB RAM) |
| Boot Disk | Ubuntu 22.04 LTS, 100GB SSD |
| Firewall | Allow HTTP, HTTPS |

4. Click **Create** (takes 1-2 minutes)

#### 1.3 SSH into VM

1. Click **SSH** button next to your VM
2. A terminal window opens in browser

#### 1.4 Run Setup Script

```bash
# Download and run setup script (ONE command)
# Use tr -d '\r' to fix Windows line endings
curl -sSL https://raw.githubusercontent.com/vinayaklearnsML2022/voters_project/main/cloud/setup.sh | tr -d '\r' | bash -s -- --port 8052
```

**Custom port:** Change `8052` to any port you want.

**What this installs:**
- Python 3.10+
- Tesseract OCR with Tamil language pack
- All Python dependencies (pymupdf, pytesseract, pillow, openpyxl, streamlit)
- Starts web UI automatically

#### 1.5 Setup Notifications (Phone)

1. Install **Ntfy** app on your phone
   - Android: [Play Store](https://play.google.com/store/apps/details?id=io.heckel.ntfy)
   - iOS: [App Store](https://apps.apple.com/app/ntfy/id1625396347)

2. Open app → Tap **+** → Enter topic: `vinayak-voter-alerts`

3. Subscribe

---

### Phase 2: Daily Usage (No Terminal Needed!)

#### 2.1 Start Streamlit

```bash
cd ~/voter_analytics && source venv/bin/activate && nohup streamlit run cloud/voter_processor_ui.py --server.port 8052 --server.address 0.0.0.0 > ~/streamlit.log 2>&1 &
```

#### 2.2 Access Web UI

Open browser and go to:
```
http://YOUR_VM_IP:8052
```

You'll see:

```
┌─────────────────────────────────────────────────────────────┐
│                                                             │
│   🗳️ VOTER ANALYTICS PROCESSOR                             │
│   ─────────────────────────────────────────                │
│                                                             │
│   📁 UPLOAD CONSTITUENCY                                    │
│   ┌─────────────────────────────────────────┐              │
│   │                                         │              │
│   │     Drag & Drop PDF files here          │              │
│   │           or click to browse            │              │
│   │                                         │              │
│   └─────────────────────────────────────────┘              │
│                                                             │
│   Constituency Name: [_______________________]              │
│                                                             │
│   [🚀 Start Processing]                                     │
│                                                             │
└─────────────────────────────────────────────────────────────┘
```

#### 2.2 Upload PDFs

1. Enter **Constituency Name** (e.g., `1-Gummidipoondi`)
2. **Drag & Drop** all PDF files for that constituency
3. Click **🚀 Start Processing**

#### 2.3 Processing Begins

The UI shows live progress:

```
┌─────────────────────────────────────────────────────────────┐
│                                                             │
│   📊 PROCESSING STATUS                                      │
│   ─────────────────────                                    │
│                                                             │
│   Constituency: 1-Gummidipoondi                            │
│   Status: 🔄 Processing                                     │
│                                                             │
│   ┌─────────────────────────────────────────┐              │
│   │ Phase 1/4: Extracting cards from PDFs   │              │
│   │ ████████████░░░░░░░░ 60%                │              │
│   │ 27/45 PDFs processed                    │              │
│   └─────────────────────────────────────────┘              │
│                                                             │
│   📈 Statistics                                             │
│   ├─ PDFs Processed: 27/45                                 │
│   ├─ Cards Extracted: 24,350                               │
│   ├─ Time Elapsed: 45m 23s                                 │
│   └─ Estimated Remaining: ~30 minutes                      │
│                                                             │
└─────────────────────────────────────────────────────────────┘
```

**You can close the browser!** Processing continues on VM.

#### 2.4 Get Notification

When processing completes, you receive a push notification:

```
┌─────────────────────────────────┐
│ 🔔 Ntfy                    now  │
├─────────────────────────────────┤
│ ✅ Processing Complete!         │
│                                 │
│ Constituency: 1-Gummidipoondi   │
│ Total Cards: 45,230             │
│ Missing Age: 234 (0.5%)         │
│ Missing Gender: 189 (0.4%)      │
│ Time: 2h 15m                    │
│                                 │
│ Excel ready for download!       │
└─────────────────────────────────┘
```

#### 2.5 Download Excel

1. Open Web UI
2. Go to **📥 Downloads** section
3. Click **Download Excel**

```
┌─────────────────────────────────────────────────────────────┐
│                                                             │
│   📥 DOWNLOAD RESULTS                                       │
│   ───────────────────                                      │
│                                                             │
│   ┌─────────────────────────────────────────────────────┐  │
│   │ File                        │ Size   │ Action       │  │
│   ├─────────────────────────────┼────────┼──────────────┤  │
│   │ 1-Gummidipoondi_excel.xlsx  │ 4.2 MB │ [Download]   │  │
│   │ 2-Ponneri_excel.xlsx        │ 3.8 MB │ [Download]   │  │
│   │ 3-Tiruvallur_excel.xlsx     │ 4.5 MB │ [Download]   │  │
│   └─────────────────────────────┴────────┴──────────────┘  │
│                                                             │
│   [📦 Download All as ZIP]    [🗑️ Clear Old Files]         │
│                                                             │
└─────────────────────────────────────────────────────────────┘
```

#### 2.6 Auto Cleanup

After downloading:
- Temp card images are **automatically deleted**
- Disk space is freed for next constituency
- Only Excel files are kept until you clear them

---

## Excel Output Format

The generated Excel contains:

| Column | Description | Example |
|--------|-------------|---------|
| S.No | Serial number | 1, 2, 3... |
| Part No. | PDF part number | 1, 2, 3... |
| Voter ID | Unique voter ID | ABC1234567 |
| Name | Voter name (Tamil) | முருகன் |
| Relation Type | Father/Husband/Mother | Father |
| Relation Name | Relation's name | செல்வம் |
| House No | House number | 123/A |
| Age | Voter's age | 45 |
| Gender | Male/Female/Third Gender | Male |
| Constituency | Constituency name | 1-Gummidipoondi |
| Source Folder | PDF folder name | TAM-1-WI... |
| Card File | Image filename | 1.png |

**Missing data is highlighted in yellow** for easy identification.

---

## Processing Pipeline Details

### Phase 1: PDF to Images (Fastest)

```
Input: 45 PDF files
       ↓
Extract pages (skip first 3 + last 1)
       ↓
Divide each page into 3×10 grid = 30 cards/page
       ↓
Save as PNG (low compression for speed)
       ↓
Output: ~40,000 card images
```

**Time:** ~15-20 minutes for 45 PDFs

### Phase 2: OCR Processing (Slowest)

```
Input: ~40,000 card images
       ↓
For each image:
  → Open image
  → Run Tesseract OCR (Tamil + English)
  → Extract: Voter ID, Name, Age, Gender, etc.
  → Save to memory
       ↓
Progress: Updates every 50 cards
Checkpoint: Saves every 200 cards (resume if crash)
       ↓
Output: Structured data for all cards
```

**Time:** ~1.5-2 hours for 40,000 cards

### Phase 3: Fix Missing Age/Gender

```
Input: Cards with missing Age or Gender
       ↓
For each missing card:
  → Crop bottom 30% of image (where Age/Gender appears)
  → Try multiple preprocessing (contrast, binarize, etc.)
  → Re-run OCR
  → Update data if found
       ↓
Output: Improved data with fewer missing values
```

**Time:** ~15-30 minutes

### Phase 4: Generate Excel

```
Input: All extracted data
       ↓
Create Excel workbook
  → Add headers
  → Write all rows
  → Highlight missing cells in yellow
  → Set column widths
       ↓
Save Excel file
Delete temp images (auto-cleanup)
Send notification
       ↓
Output: Final Excel file ready for download
```

**Time:** ~2-3 minutes

---

## Queue System (Multiple Constituencies)

Upload multiple constituencies - they process one by one:

```
┌─────────────────────────────────────────────────────────────┐
│                                                             │
│   📋 PROCESSING QUEUE                                       │
│   ──────────────────                                       │
│                                                             │
│   1. ✅ 1-Gummidipoondi      Complete    [Download]        │
│   2. ✅ 2-Ponneri            Complete    [Download]        │
│   3. 🔄 3-Tiruvallur         Processing  45%               │
│   4. ⏳ 4-Ambattur           Queued      --                │
│   5. ⏳ 5-Madhavaram         Queued      --                │
│                                                             │
│   [➕ Add More]  [⏸️ Pause Queue]  [🗑️ Clear Completed]     │
│                                                             │
└─────────────────────────────────────────────────────────────┘
```

---

## Error Handling

### If Processing Fails

1. You receive error notification:
   ```
   ❌ Processing Error
   Constituency: 3-Tiruvallur
   Error: Out of memory at PDF #23
   ```

2. Check logs in Web UI

3. **Resume from checkpoint** - progress is saved!
   - Click "Resume" button
   - Processing continues from where it stopped

### If VM Restarts

- Checkpoints saved every 200 cards
- On restart, run: `python3 resume_processing.py`
- Continues automatically

---

## Cost Summary

### GCP Costs (With Free Credits)

| Item | Cost | Your Cost |
|------|------|-----------|
| e2-standard-8 VM | ~₹22/hour | **FREE** (using $300 credits) |
| 100GB SSD | Included | **FREE** |
| Network | ~₹1/GB | Minimal |

**For 234 constituencies:**
- ~585 hours of processing
- ~₹12,870 in VM costs
- **Covered by $300 free credits!** ✅

### After Free Credits

If you need more:
- Same VM costs ~₹22/hour
- Or use 3 smaller VMs in parallel to save time

---

## Scaling to 3 VMs

When ready to scale:

```
VM 1: Constituencies 1-78      (runs independently)
VM 2: Constituencies 79-156    (runs independently)
VM 3: Constituencies 157-234   (runs independently)
```

Each VM:
- Has its own Web UI
- Processes its own queue
- Sends notifications to same phone
- Reduces total time from 25 days → 8-9 days

---

## File Structure on VM

```
/home/user/
├── voter_analytics/
│   ├── voter_processor_ui.py      # Streamlit Web UI
│   ├── process_batch_headless.py  # CLI processor
│   ├── setup.sh                   # Setup script
│   │
│   ├── uploads/                   # Uploaded PDFs (temp)
│   │   └── 1-Gummidipoondi/
│   │       ├── TAM-1-WI.pdf
│   │       ├── TAM-2-WI.pdf
│   │       └── ...
│   │
│   ├── processing/                # Temp card images
│   │   └── .1-Gummidipoondi_temp_cards/
│   │       ├── TAM-1-WI/
│   │       │   ├── 1.png
│   │       │   ├── 2.png
│   │       │   └── ...
│   │       └── ...
│   │
│   ├── output/                    # Final Excel files
│   │   ├── 1-Gummidipoondi_excel.xlsx
│   │   ├── 2-Ponneri_excel.xlsx
│   │   └── ...
│   │
│   └── logs/                      # Processing logs
│       └── processing.log
```

---

## Quick Reference Commands

### Start Web UI (if stopped)
```bash
cd ~/voter_analytics && source venv/bin/activate
nohup streamlit run cloud/voter_processor_ui.py --server.port 8052 --server.address 0.0.0.0 > ~/streamlit.log 2>&1 &
```

### Kill Streamlit
```bash
pkill -f streamlit
```

### Reset Stuck Status
```bash
echo '{"processing": false, "current_constituency": null, "pid": null, "queue": [], "completed": [], "errors": []}' > ~/voter_analytics/.processing_status.json
```

### Check Processing Status
```bash
tail -f ~/voter_analytics/logs/processing.log
```

### Manual Test Notification
```bash
curl -d "Test notification!" ntfy.sh/voter-analytics-YOUR-SECRET
```

### Check Disk Space
```bash
df -h
```

### Clear Old Temp Files
```bash
rm -rf ~/voter_analytics/processing/*
```

---

## Troubleshooting

| Problem | Solution |
|---------|----------|
| Web UI not loading | Check if VM is running, check firewall |
| OCR missing data | Normal - use Missing Data Finder tool |
| Processing slow | Check CPU usage, RAM usage |
| Out of disk space | Clear old temp files |
| No notification | Check Ntfy app subscription |
| VM crashed | SSH in, run resume script |

---

## Support

- Check logs: `~/voter_analytics/logs/processing.log`
- Test notification: `curl -d "test" ntfy.sh/your-topic`
- Resume processing: `python3 resume_processing.py`
