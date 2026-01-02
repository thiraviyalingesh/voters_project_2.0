# Windows VM Setup Guide for Voter Analytics

## Why Windows VM?

| Factor | Linux (Ubuntu) | Windows Server |
|--------|---------------|----------------|
| Cost | ~$0.34/hour | ~$0.50-0.60/hour |
| License | Free | **Included in price** |
| Multi-worker OCR | Hangs | **Works** |
| Setup | Script | Manual |

Windows handles Python multiprocessing correctly with Tesseract, so 8 workers will work.

---

## Step 1: Create Windows VM in GCP

1. Go to **Compute Engine → VM instances**
2. Click **Create Instance**
3. Configure:

| Setting | Value |
|---------|-------|
| Name | `voter-analytics-windows` |
| Region | `asia-south1 (Mumbai)` |
| Machine Type | `e2-standard-8` (8 vCPU, 32GB RAM) |
| Boot Disk | Click **Change** |
| - Operating system | **Windows Server** |
| - Version | **Windows Server 2022 Datacenter** |
| - Size | **100 GB SSD** |
| Firewall | Allow HTTP, HTTPS |

4. Click **Create** (takes 2-3 minutes)

---

## Step 2: Connect to VM

1. Wait for VM status to show **Running**
2. Click **Set Windows password** → Set username and password
3. Click **RDP** → **Download the RDP file**
4. Open the RDP file → Enter password → Connect

---

## Step 3: Install Python

Open **PowerShell as Administrator** and run:

```powershell
# Install Python via winget
winget install Python.Python.3.11

# Close and reopen PowerShell, then verify:
python --version
```

---

## Step 4: Install Tesseract OCR

1. Download installer from: https://github.com/UB-Mannheim/tesseract/wiki
   - Direct link: https://digi.bib.uni-mannheim.de/tesseract/tesseract-ocr-w64-setup-5.3.3.20231005.exe

2. Run the installer:
   - Click **Next** through the wizard
   - On **Select Components**, expand **Additional language data**
   - Check **Tamil** and **English**
   - Complete installation

3. Add Tesseract to PATH:
   ```powershell
   # Run in PowerShell as Admin
   [Environment]::SetEnvironmentVariable("Path", $env:Path + ";C:\Program Files\Tesseract-OCR", "Machine")
   ```

4. Restart PowerShell and verify:
   ```powershell
   tesseract --version
   tesseract --list-langs
   ```

---

## Step 5: Install Python Packages

```powershell
pip install pymupdf pytesseract pillow openpyxl streamlit requests
```

---

## Step 6: Clone Your Repository

```powershell
cd C:\
git clone https://github.com/vinayaklearnsML2022/voters_project.git
cd voters_project
```

---

## Step 7: Run the Processor

### Option A: GUI Mode (voter_counter_app_fast_4.0.py)

```powershell
python voter_counter_app_fast_4.0.py
```

- Select constituency folder
- Click Start
- Uses 8 workers automatically

### Option B: Headless Mode (for background processing)

```powershell
python cloud/process_batch_headless.py C:\path\to\constituency --ntfy-topic vinayak-voter-alerts --workers 8
```

---

## Step 8: Open Firewall for Streamlit (Optional)

If you want to access Web UI from your PC:

1. In Windows VM, open **Windows Defender Firewall**
2. Click **Advanced settings**
3. **Inbound Rules → New Rule**
4. Port → TCP → 8501 → Allow → Name: "Streamlit"

5. In GCP Console:
   - VPC Network → Firewall → Create Rule
   - Allow TCP 8501 from 0.0.0.0/0

6. Run Streamlit:
   ```powershell
   cd C:\voters_project\cloud
   streamlit run voter_processor_ui.py --server.port 8501 --server.address 0.0.0.0
   ```

7. Access from your PC: `http://VM_EXTERNAL_IP:8501`

---

## Cost Estimate

| Item | Cost |
|------|------|
| Windows VM (e2-standard-8) | ~$0.55/hour |
| 234 constituencies × 2.5 hours | 585 hours |
| **Total** | ~$320 |
| **GCP Free Credits** | $300 |
| **Extra needed** | ~$20-50 |

---

## Tips

1. **Stop VM when not using** - You only pay when running
2. **Use RDP to connect** - Not SSH like Linux
3. **8 workers will work** - Windows multiprocessing is reliable
4. **Upload PDFs via RDP** - Drag & drop files into RDP window

---

## Troubleshooting

| Problem | Solution |
|---------|----------|
| Python not found | Restart PowerShell after install |
| Tesseract not found | Add to PATH and restart PowerShell |
| RDP won't connect | Check VM is running, check firewall |
| Slow RDP | Use lower resolution in RDP settings |
