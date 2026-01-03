# Voter Analytics

Extract voter data from Tamil Nadu Electoral Roll PDFs using OCR.

## Quick Start (Cloud VM)

```bash
curl -sSL https://raw.githubusercontent.com/vinayaklearnsML2022/voters_project/main/cloud/setup.sh | tr -d '\r' | bash -s -- --port 8052
```

Then start Streamlit:
```bash
cd ~/voter_analytics && source venv/bin/activate && nohup streamlit run cloud/voter_processor_ui.py --server.port 8052 --server.address 0.0.0.0 > ~/streamlit.log 2>&1 &
```

Access at: `http://<VM_IP>:8052`

## Documentation

| Guide | Description |
|-------|-------------|
| [VM Setup Guide](docs/vm_setup_guide.md) | Quick setup commands and troubleshooting |
| [Complete System Guide](docs/complete_system_guide.md) | Full walkthrough with architecture |
| [Windows VM Setup](docs/windows_vm_setup.md) | Setup on Windows Server |
| [Multiple VM Setup](docs/multiple_vm_setup_guide.md) | Scaling to multiple VMs |

## Features

- Web UI for uploading PDFs and monitoring progress
- OCR with Tamil language support (Tesseract)
- Checkpoint/resume capability
- Push notifications via Ntfy
- Excel output with highlighted missing data

## Processing Pipeline

1. **Phase 1:** Extract voter cards from PDFs
2. **Phase 2:** OCR all cards (Tamil + English)
3. **Phase 3:** Fix missing Age/Gender with enhanced OCR
4. **Phase 4:** Generate Excel with formatting

## Common Commands

```bash
# Start Streamlit
cd ~/voter_analytics && source venv/bin/activate
streamlit run cloud/voter_processor_ui.py --server.port 8052 --server.address 0.0.0.0

# Reset stuck status
echo '{"processing": false, "current_constituency": null, "pid": null, "queue": [], "completed": [], "errors": []}' > ~/voter_analytics/.processing_status.json

# Kill Streamlit
pkill -f streamlit

# Check logs
tail -f ~/voter_analytics/logs/*.log

# Update code
cd ~/voter_analytics && git pull origin main
```

## GCP Firewall

```bash
gcloud config set project YOUR_PROJECT_ID
gcloud compute firewall-rules create allow-streamlit-8052 --allow tcp:8052 --direction INGRESS
```
