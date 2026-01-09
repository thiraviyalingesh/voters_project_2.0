# Start/Stop Voter Analytics 2.0

## First Time Setup (Fresh VM)

```bash
curl -sSL https://raw.githubusercontent.com/thiraviyalingesh/voters_project_2.0/main/cloud/setup.sh | bash
```

---

## Start VM and Streamlit

### 1. Start VM
```bash
gcloud compute instances start YOUR_INSTANCE_NAME --zone=YOUR_ZONE
```

### 2. SSH into VM
```bash
gcloud compute ssh YOUR_INSTANCE_NAME --zone=YOUR_ZONE
```

### 3. Start Streamlit
```bash
cd ~/voter_analytics_2.0 && source venv/bin/activate && nohup streamlit run cloud/voter_processor_ui.py --server.port 8053 --server.address 0.0.0.0 > ~/streamlit_2.0.log 2>&1 &
```

### 4. Verify Running
```bash
ps aux | grep streamlit
```

### 5. Access Web UI
```
http://<VM_EXTERNAL_IP>:8053
```

---

## Stop VM

### 1. Kill Streamlit (optional - VM stop will kill it anyway)
```bash
pkill -f streamlit
```

### 2. Stop VM
```bash
gcloud compute instances stop YOUR_INSTANCE_NAME --zone=YOUR_ZONE
```

---

## Reset Stuck Status

If status shows processing but nothing is running:
```bash
echo '{"processing": false, "current_constituency": null, "pid": null, "queue": [], "completed": [], "errors": []}' > ~/voter_analytics_2.0/.processing_status.json
```

---

## Notes

- Your data, uploads, and code are preserved when VM stops
- Only running processes stop
- Processing can resume from checkpoints if interrupted
- Port: 8053 (different from boss's 8052)
