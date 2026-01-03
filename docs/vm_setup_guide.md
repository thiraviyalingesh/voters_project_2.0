# Voter Analytics - VM Setup Guide

## Quick Setup (Fresh Ubuntu 22.04 VM)

### One Command Setup

```bash
curl -sSL https://raw.githubusercontent.com/vinayaklearnsML2022/voters_project/main/cloud/setup.sh | tr -d '\r' | bash -s -- --port 8052
```

**Note:** The `tr -d '\r'` fixes Windows line endings if present.

### Custom Port

```bash
# Default port 8501
curl -sSL https://raw.githubusercontent.com/vinayaklearnsML2022/voters_project/main/cloud/setup.sh | tr -d '\r' | bash

# Custom port (e.g., 8052)
curl -sSL https://raw.githubusercontent.com/vinayaklearnsML2022/voters_project/main/cloud/setup.sh | tr -d '\r' | bash -s -- --port 8052
```

---

## Start Streamlit

```bash
cd ~/voter_analytics && source venv/bin/activate && nohup streamlit run cloud/voter_processor_ui.py --server.port 8052 --server.address 0.0.0.0 > ~/streamlit.log 2>&1 &
```

### Verify Running

```bash
ps aux | grep streamlit
```

### Access Web UI

```
http://<VM_EXTERNAL_IP>:8052
```

Get your VM's external IP:
```bash
curl -s ifconfig.me
```

---

## GCP Firewall Setup

### Set Project

```bash
gcloud config set project YOUR_PROJECT_ID
```

### Create Firewall Rule

```bash
gcloud compute firewall-rules create allow-streamlit-8052 --allow tcp:8052 --direction INGRESS --priority 1000
```

### Or via GCP Console

1. Go to **VPC Network -> Firewall**
2. Click **Create Firewall Rule**
3. Name: `allow-streamlit-8052`
4. Direction: **Ingress**
5. Targets: **All instances**
6. Source IP ranges: `0.0.0.0/0`
7. Protocols and ports: **tcp:8052**
8. Click **Create**

---

## Common Commands

### Reset Status (Fix Stuck Processing)

```bash
echo '{"processing": false, "current_constituency": null, "pid": null, "queue": [], "completed": [], "errors": []}' > ~/voter_analytics/.processing_status.json
```

### Kill Streamlit

```bash
pkill -f streamlit
```

### Restart Streamlit

```bash
pkill -f streamlit
cd ~/voter_analytics && source venv/bin/activate && nohup streamlit run cloud/voter_processor_ui.py --server.port 8052 --server.address 0.0.0.0 > ~/streamlit.log 2>&1 &
```

### Check Logs

```bash
tail -f ~/streamlit.log
tail -f ~/voter_analytics/logs/*.log
```

### Update Code

```bash
cd ~/voter_analytics && git pull origin main
```

---

## Using Systemd Service (Auto-start)

### Start Service

```bash
sudo systemctl start voter-analytics
sudo systemctl enable voter-analytics
```

### Check Status

```bash
sudo systemctl status voter-analytics
```

### Stop Service

```bash
sudo systemctl stop voter-analytics
```

---

## Troubleshooting

### Port Already in Use

```bash
pkill -f streamlit
# Or kill specific port
fuser -k 8052/tcp
```

### Process Crashed But Status Shows Running

Use the **Reset Stuck Status** button in the Streamlit UI, or run:
```bash
echo '{"processing": false, "current_constituency": null, "pid": null, "queue": [], "completed": [], "errors": []}' > ~/voter_analytics/.processing_status.json
```

### Connection Timeout

1. Check if Streamlit is running: `ps aux | grep streamlit`
2. Check firewall rule exists: `gcloud compute firewall-rules list`
3. Verify port is correct in both Streamlit command and firewall rule

### Notification Not Working

1. Enter your Ntfy topic in the sidebar
2. Click **Send Test** to verify
3. Check topic name is correct (no spaces, valid characters)
