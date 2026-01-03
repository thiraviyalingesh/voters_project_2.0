# Start/Stop Voter Analytics

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
cd ~/voter_analytics && source venv/bin/activate && nohup streamlit run cloud/voter_processor_ui.py --server.port 8052 --server.address 0.0.0.0 > ~/streamlit.log 2>&1 &
```

### 4. Verify Running
```bash
ps aux | grep streamlit
```

### 5. Access Web UI
```
http://<VM_EXTERNAL_IP>:8052
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
echo '{"processing": false, "current_constituency": null, "pid": null, "queue": [], "completed": [], "errors": []}' > ~/voter_analytics/.processing_status.json
```

---

## Notes

- Your data, uploads, and code are preserved when VM stops
- Only running processes stop
- Processing can resume from checkpoints if interrupted
