"""
Voter Analytics - Streamlit Web UI
Upload PDFs, monitor processing, download results.

Usage:
    streamlit run voter_processor_ui.py --server.port 8501 --server.address 0.0.0.0
"""

import streamlit as st
import os
import sys
import time
import json
import subprocess
import threading
from pathlib import Path
from datetime import datetime
import requests

# ============== CONFIGURATION ==============
# Edit these settings as needed
NTFY_TOPIC = "voter-analytics-alerts"  # Your Ntfy topic
BASE_DIR = Path.home() / "voter_analytics"
UPLOAD_DIR = BASE_DIR / "uploads"
OUTPUT_DIR = BASE_DIR / "output"
PROCESSING_DIR = BASE_DIR / "processing"
LOGS_DIR = BASE_DIR / "logs"

# Create directories if they don't exist
for d in [UPLOAD_DIR, OUTPUT_DIR, PROCESSING_DIR, LOGS_DIR]:
    d.mkdir(parents=True, exist_ok=True)

# ============== STATE MANAGEMENT ==============

def get_status_file():
    """Get path to status file."""
    return BASE_DIR / ".processing_status.json"

def load_status():
    """Load processing status from file."""
    status_file = get_status_file()
    if status_file.exists():
        try:
            with open(status_file, 'r') as f:
                return json.load(f)
        except:
            pass
    return {
        'processing': False,
        'current_constituency': None,
        'phase': 0,
        'progress': 0,
        'total': 0,
        'start_time': None,
        'queue': [],
        'completed': [],
        'errors': []
    }

def save_status(status):
    """Save processing status to file."""
    with open(get_status_file(), 'w') as f:
        json.dump(status, f)

def add_to_queue(constituency_name, folder_path):
    """Add constituency to processing queue."""
    status = load_status()
    status['queue'].append({
        'name': constituency_name,
        'folder': str(folder_path),
        'added_at': datetime.now().isoformat()
    })
    save_status(status)

def get_completed_files():
    """Get list of completed Excel files."""
    files = []
    if OUTPUT_DIR.exists():
        for f in OUTPUT_DIR.glob("*.xlsx"):
            files.append({
                'name': f.name,
                'path': str(f),
                'size': f.stat().st_size,
                'modified': datetime.fromtimestamp(f.stat().st_mtime)
            })
    return sorted(files, key=lambda x: x['modified'], reverse=True)

def send_notification(title, message):
    """Send push notification via Ntfy."""
    if NTFY_TOPIC:
        try:
            requests.post(
                f"https://ntfy.sh/{NTFY_TOPIC}",
                headers={"Title": title},
                data=message.encode('utf-8'),
                timeout=10
            )
        except:
            pass

def run_processor(folder_path, constituency_name):
    """Run the headless processor in background."""
    status = load_status()
    status['processing'] = True
    status['current_constituency'] = constituency_name
    status['phase'] = 1
    status['progress'] = 0
    status['start_time'] = datetime.now().isoformat()
    save_status(status)

    log_file = LOGS_DIR / f"{constituency_name}_{datetime.now().strftime('%Y%m%d_%H%M%S')}.log"

    try:
        # Run processor
        cmd = [
            sys.executable,
            str(BASE_DIR / "process_batch_headless.py"),
            str(folder_path),
            "--ntfy-topic", NTFY_TOPIC
        ]

        with open(log_file, 'w') as log:
            process = subprocess.Popen(
                cmd,
                stdout=log,
                stderr=subprocess.STDOUT,
                cwd=str(BASE_DIR)
            )
            process.wait()

        # Update status
        status = load_status()
        if process.returncode == 0:
            status['completed'].append({
                'name': constituency_name,
                'completed_at': datetime.now().isoformat()
            })
        else:
            status['errors'].append({
                'name': constituency_name,
                'error': f"Process exited with code {process.returncode}",
                'time': datetime.now().isoformat()
            })
            send_notification(
                "❌ Processing Error",
                f"Constituency: {constituency_name}\nCheck logs for details"
            )
    except Exception as e:
        status = load_status()
        status['errors'].append({
            'name': constituency_name,
            'error': str(e),
            'time': datetime.now().isoformat()
        })
        send_notification("❌ Processing Error", f"Constituency: {constituency_name}\nError: {str(e)}")

    # Check if more in queue
    status = load_status()
    status['processing'] = False
    status['current_constituency'] = None

    if status['queue']:
        next_item = status['queue'].pop(0)
        save_status(status)
        run_processor(next_item['folder'], next_item['name'])
    else:
        save_status(status)

def start_processing(folder_path, constituency_name):
    """Start processing in background thread."""
    thread = threading.Thread(
        target=run_processor,
        args=(folder_path, constituency_name),
        daemon=True
    )
    thread.start()

def format_size(size_bytes):
    """Format file size."""
    for unit in ['B', 'KB', 'MB', 'GB']:
        if size_bytes < 1024:
            return f"{size_bytes:.1f} {unit}"
        size_bytes /= 1024
    return f"{size_bytes:.1f} TB"

def format_time_ago(dt):
    """Format datetime as 'X ago'."""
    if isinstance(dt, str):
        dt = datetime.fromisoformat(dt)
    diff = datetime.now() - dt
    seconds = diff.total_seconds()

    if seconds < 60:
        return "just now"
    elif seconds < 3600:
        mins = int(seconds / 60)
        return f"{mins}m ago"
    elif seconds < 86400:
        hours = int(seconds / 3600)
        return f"{hours}h ago"
    else:
        days = int(seconds / 86400)
        return f"{days}d ago"

# ============== STREAMLIT UI ==============

st.set_page_config(
    page_title="Voter Analytics Processor",
    page_icon="🗳️",
    layout="wide"
)

# Custom CSS
st.markdown("""
<style>
    .big-font {
        font-size: 24px !important;
        font-weight: bold;
    }
    .status-box {
        padding: 20px;
        border-radius: 10px;
        margin: 10px 0;
    }
    .processing {
        background-color: #fff3cd;
        border: 1px solid #ffc107;
    }
    .completed {
        background-color: #d4edda;
        border: 1px solid #28a745;
    }
    .error {
        background-color: #f8d7da;
        border: 1px solid #dc3545;
    }
    .stProgress > div > div > div > div {
        background-color: #28a745;
    }
</style>
""", unsafe_allow_html=True)

# Header
st.title("🗳️ Voter Analytics Processor")
st.markdown("---")

# Sidebar - Configuration
with st.sidebar:
    st.header("⚙️ Settings")

    new_topic = st.text_input("Ntfy Topic", value=NTFY_TOPIC)
    if new_topic != NTFY_TOPIC:
        st.info("Restart app to apply new topic")

    st.markdown("---")

    st.subheader("📱 Test Notification")
    if st.button("Send Test"):
        send_notification("🔔 Test", "Notification test from Voter Analytics!")
        st.success("Sent!")

    st.markdown("---")

    st.subheader("📊 System Info")
    import multiprocessing
    st.text(f"CPU Cores: {multiprocessing.cpu_count()}")

    # Disk space
    import shutil
    total, used, free = shutil.disk_usage(BASE_DIR)
    st.text(f"Disk Free: {format_size(free)}")

# Main content - Tabs
tab1, tab2, tab3, tab4 = st.tabs(["📁 Upload", "📊 Status", "📥 Downloads", "📋 Logs"])

# ============== TAB 1: Upload ==============
with tab1:
    st.header("Upload Constituency PDFs")

    col1, col2 = st.columns([2, 1])

    with col1:
        constituency_name = st.text_input(
            "Constituency Name",
            placeholder="e.g., 1-Gummidipoondi",
            help="This will be used as the folder name and in the output file"
        )

    with col2:
        st.markdown("<br>", unsafe_allow_html=True)

    uploaded_files = st.file_uploader(
        "Upload PDF Files",
        type=['pdf'],
        accept_multiple_files=True,
        help="Upload all PDF files for this constituency"
    )

    if uploaded_files and constituency_name:
        st.info(f"📎 {len(uploaded_files)} PDF files selected")

        if st.button("🚀 Start Processing", type="primary"):
            # Create constituency folder
            constituency_folder = UPLOAD_DIR / constituency_name
            constituency_folder.mkdir(parents=True, exist_ok=True)

            # Save uploaded files
            progress_bar = st.progress(0)
            status_text = st.empty()

            for i, file in enumerate(uploaded_files):
                status_text.text(f"Saving {file.name}...")
                file_path = constituency_folder / file.name
                with open(file_path, 'wb') as f:
                    f.write(file.getbuffer())
                progress_bar.progress((i + 1) / len(uploaded_files))

            status_text.text("Files saved! Starting processing...")

            # Check if already processing
            status = load_status()
            if status['processing']:
                # Add to queue
                add_to_queue(constituency_name, constituency_folder)
                st.success(f"✅ Added to queue. Position: {len(status['queue']) + 1}")
            else:
                # Start immediately
                start_processing(constituency_folder, constituency_name)
                st.success("🚀 Processing started!")

            time.sleep(1)
            st.rerun()

    elif uploaded_files and not constituency_name:
        st.warning("⚠️ Please enter a constituency name")

# ============== TAB 2: Status ==============
with tab2:
    st.header("Processing Status")

    status = load_status()

    # Current processing
    if status['processing'] and status['current_constituency']:
        st.markdown(f"""
        <div class="status-box processing">
            <h3>🔄 Currently Processing</h3>
            <p><strong>Constituency:</strong> {status['current_constituency']}</p>
            <p><strong>Started:</strong> {format_time_ago(status['start_time']) if status['start_time'] else 'Unknown'}</p>
        </div>
        """, unsafe_allow_html=True)

        # Check for checkpoint file to get progress
        checkpoint_file = UPLOAD_DIR.parent / f".{status['current_constituency']}_checkpoint.json"
        if checkpoint_file.exists():
            try:
                with open(checkpoint_file, 'r') as f:
                    cp = json.load(f)
                phase = cp.get('phase', 0)
                st.progress(phase / 4, text=f"Phase {phase}/4")
            except:
                st.progress(0.25, text="Processing...")
        else:
            st.progress(0.1, text="Starting...")

        if st.button("🔄 Refresh"):
            st.rerun()
    else:
        st.info("No active processing")

    # Queue
    st.subheader("📋 Queue")
    if status['queue']:
        for i, item in enumerate(status['queue'], 1):
            st.text(f"  {i}. {item['name']} (added {format_time_ago(item['added_at'])})")
    else:
        st.text("  Queue is empty")

    # Recent completions
    st.subheader("✅ Recently Completed")
    if status['completed']:
        for item in status['completed'][-5:]:
            st.success(f"  {item['name']} - {format_time_ago(item['completed_at'])}")
    else:
        st.text("  No completed jobs yet")

    # Errors
    if status['errors']:
        st.subheader("❌ Errors")
        for item in status['errors'][-3:]:
            st.error(f"  {item['name']}: {item['error']}")

        if st.button("Clear Errors"):
            status['errors'] = []
            save_status(status)
            st.rerun()

# ============== TAB 3: Downloads ==============
with tab3:
    st.header("Download Results")

    files = get_completed_files()

    if files:
        for f in files:
            col1, col2, col3, col4 = st.columns([3, 1, 1, 1])

            with col1:
                st.text(f["name"])
            with col2:
                st.text(format_size(f["size"]))
            with col3:
                st.text(format_time_ago(f["modified"]))
            with col4:
                with open(f["path"], "rb") as fp:
                    st.download_button(
                        "📥",
                        data=fp,
                        file_name=f["name"],
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        key=f["name"]
                    )

        st.markdown("---")

        # Bulk actions
        col1, col2 = st.columns(2)
        with col1:
            if st.button("🗑️ Clear Old Files"):
                # Keep only last 10 files
                for f in files[10:]:
                    try:
                        os.remove(f["path"])
                    except:
                        pass
                st.success("Cleaned up old files")
                st.rerun()
    else:
        st.info("No completed files yet")

    # Disk usage
    st.markdown("---")
    st.subheader("💾 Storage")

    import shutil
    total, used, free = shutil.disk_usage(BASE_DIR)

    col1, col2, col3 = st.columns(3)
    col1.metric("Total", format_size(total))
    col2.metric("Used", format_size(used))
    col3.metric("Free", format_size(free))

    # Usage bar
    usage_pct = used / total
    st.progress(usage_pct, text=f"{usage_pct*100:.1f}% used")

# ============== TAB 4: Logs ==============
with tab4:
    st.header("Processing Logs")

    log_files = sorted(LOGS_DIR.glob("*.log"), key=lambda x: x.stat().st_mtime, reverse=True)

    if log_files:
        selected_log = st.selectbox(
            "Select Log File",
            options=log_files,
            format_func=lambda x: f"{x.name} ({format_time_ago(datetime.fromtimestamp(x.stat().st_mtime))})"
        )

        if selected_log:
            # Show last 100 lines
            with open(selected_log, 'r') as f:
                lines = f.readlines()

            st.text_area(
                "Log Output",
                value="".join(lines[-100:]),
                height=400,
                disabled=True
            )

            col1, col2 = st.columns(2)
            with col1:
                if st.button("🔄 Refresh Log"):
                    st.rerun()
            with col2:
                with open(selected_log, 'r') as f:
                    st.download_button(
                        "📥 Download Full Log",
                        data=f.read(),
                        file_name=selected_log.name,
                        mime="text/plain"
                    )
    else:
        st.info("No log files yet")

# Footer
st.markdown("---")
st.markdown(
    "<center><small>Voter Analytics Processor v1.0 | "
    f"Last refreshed: {datetime.now().strftime('%H:%M:%S')}</small></center>",
    unsafe_allow_html=True
)

# Auto-refresh when processing
status = load_status()
if status['processing']:
    time.sleep(5)
    st.rerun()
