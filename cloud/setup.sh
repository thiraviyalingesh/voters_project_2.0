#!/bin/bash
#
# Voter Analytics - One-Time Setup Script
# Run this on a fresh Ubuntu 22.04 VM
#
# Usage: curl -sSL https://your-url/setup.sh | bash
#

set -e  # Exit on error

echo "=============================================="
echo "  Voter Analytics - Setup Script"
echo "=============================================="
echo ""

# Colors for output
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
NC='\033[0m' # No Color

print_status() {
    echo -e "${GREEN}[✓]${NC} $1"
}

print_warning() {
    echo -e "${YELLOW}[!]${NC} $1"
}

print_error() {
    echo -e "${RED}[✗]${NC} $1"
}

# Step 1: Update system
echo ""
echo "Step 1/6: Updating system packages..."
sudo apt update -qq
sudo apt upgrade -y -qq
print_status "System updated"

# Step 2: Install system dependencies
echo ""
echo "Step 2/6: Installing system dependencies..."
sudo apt install -y -qq \
    python3 \
    python3-pip \
    python3-venv \
    tesseract-ocr \
    tesseract-ocr-tam \
    tesseract-ocr-eng \
    libtesseract-dev \
    poppler-utils \
    libgl1-mesa-glx \
    libglib2.0-0
print_status "System dependencies installed"

# Step 3: Create project directory
echo ""
echo "Step 3/6: Setting up project directory..."
mkdir -p ~/voter_analytics
mkdir -p ~/voter_analytics/uploads
mkdir -p ~/voter_analytics/processing
mkdir -p ~/voter_analytics/output
mkdir -p ~/voter_analytics/logs
cd ~/voter_analytics
print_status "Project directories created"

# Step 4: Create virtual environment
echo ""
echo "Step 4/6: Creating Python virtual environment..."
python3 -m venv venv
source venv/bin/activate
print_status "Virtual environment created"

# Step 5: Install Python packages
echo ""
echo "Step 5/6: Installing Python packages..."
pip install --upgrade pip -q
pip install -q \
    pymupdf \
    pytesseract \
    pillow \
    openpyxl \
    streamlit \
    requests \
    watchdog

print_status "Python packages installed"

# Step 6: Create systemd service for auto-start
echo ""
echo "Step 6/6: Setting up auto-start service..."

# Get current user
CURRENT_USER=$(whoami)
HOME_DIR=$(eval echo ~$CURRENT_USER)

sudo tee /etc/systemd/system/voter-analytics.service > /dev/null << EOF
[Unit]
Description=Voter Analytics Web UI
After=network.target

[Service]
Type=simple
User=$CURRENT_USER
WorkingDirectory=$HOME_DIR/voter_analytics
Environment="PATH=$HOME_DIR/voter_analytics/venv/bin"
ExecStart=$HOME_DIR/voter_analytics/venv/bin/streamlit run voter_processor_ui.py --server.port 8501 --server.address 0.0.0.0
Restart=always
RestartSec=10

[Install]
WantedBy=multi-user.target
EOF

sudo systemctl daemon-reload
print_status "Auto-start service configured"

# Print summary
echo ""
echo "=============================================="
echo -e "${GREEN}  Setup Complete!${NC}"
echo "=============================================="
echo ""
echo "Next steps:"
echo ""
echo "1. Upload the Python files to ~/voter_analytics/"
echo "   - voter_processor_ui.py"
echo "   - process_batch_headless.py"
echo ""
echo "2. Configure your Ntfy topic:"
echo "   Edit voter_processor_ui.py and set NTFY_TOPIC"
echo ""
echo "3. Start the service:"
echo "   sudo systemctl start voter-analytics"
echo "   sudo systemctl enable voter-analytics"
echo ""
echo "4. Access Web UI at:"
echo "   http://YOUR_VM_IP:8501"
echo ""
echo "5. Open firewall port 8501:"
echo "   GCP Console → VPC Network → Firewall → Create Rule"
echo "   - Allow TCP port 8501 from 0.0.0.0/0"
echo ""
echo "=============================================="
echo ""

# Check Tesseract installation
echo "Verifying installation..."
echo ""
tesseract --version | head -1
echo ""
tesseract --list-langs | grep -E "tam|eng" && print_status "Tamil & English language packs installed"
echo ""
python3 --version
print_status "Setup verification complete"
