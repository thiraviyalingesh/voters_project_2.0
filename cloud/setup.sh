#!/bin/bash
#
# Voter Analytics - One-Time Setup Script
# Run this on a fresh Ubuntu 22.04 VM
#
# Usage:
#   # Default port 8501:
#   curl -sSL https://raw.githubusercontent.com/vinayaklearnsML2022/voters_project/main/cloud/setup.sh | bash
#
#   # Custom port (e.g., 8080):
#   curl -sSL https://raw.githubusercontent.com/vinayaklearnsML2022/voters_project/main/cloud/setup.sh | bash -s -- --port 8080
#

set -e  # Exit on error

# ============== CONFIGURATION ==============
# Change this to use a different port
STREAMLIT_PORT="${STREAMLIT_PORT:-8053}"

# Parse command line arguments
while [[ $# -gt 0 ]]; do
    case $1 in
        --port)
            STREAMLIT_PORT="$2"
            shift 2
            ;;
        *)
            # If it's a URL, it's the repo URL
            if [[ "$1" == http* ]]; then
                CUSTOM_REPO="$1"
            fi
            shift
            ;;
    esac
done
# ============================================

echo "=============================================="
echo "  Voter Analytics - Setup Script v2.0"
echo "=============================================="
echo ""
echo "Port: $STREAMLIT_PORT"
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
echo "Step 1/7: Updating system packages..."
sudo apt update -qq
sudo apt upgrade -y -qq
print_status "System updated"

# Step 2: Install system dependencies
echo ""
echo "Step 2/7: Installing system dependencies..."
sudo apt install -y -qq \
    git \
    curl \
    wget \
    python3 \
    python3-pip \
    python3-venv \
    python3-dev \
    tesseract-ocr \
    tesseract-ocr-tam \
    tesseract-ocr-eng \
    libtesseract-dev \
    poppler-utils \
    libgl1-mesa-glx \
    libglib2.0-0 \
    libsm6 \
    libxext6 \
    libxrender-dev \
    libfontconfig1 \
    libice6
print_status "System dependencies installed"

# Step 3: Clone or create project directory
echo ""
echo "Step 3/7: Setting up project directory..."

# Default repo URL
DEFAULT_REPO="https://github.com/thiraviyalingesh/voters_project_2.0.git"
REPO_URL="${CUSTOM_REPO:-$DEFAULT_REPO}"

echo "Cloning from: $REPO_URL"
if [ -d ~/voter_analytics_2.0 ]; then
    print_warning "Directory exists. Pulling latest..."
    cd ~/voter_analytics_2.0 && git pull origin main || true
else
    git clone "$REPO_URL" ~/voter_analytics_2.0
fi

# Create required directories
mkdir -p ~/voter_analytics_2.0/uploads
mkdir -p ~/voter_analytics_2.0/uploads/output
mkdir -p ~/voter_analytics_2.0/processing
mkdir -p ~/voter_analytics_2.0/output
mkdir -p ~/voter_analytics_2.0/logs
cd ~/voter_analytics_2.0
print_status "Project directories created"

# Step 4: Create virtual environment
echo ""
echo "Step 4/7: Creating Python virtual environment..."
python3 -m venv venv
source venv/bin/activate
print_status "Virtual environment created"

# Step 5: Install Python packages
echo ""
echo "Step 5/7: Installing Python packages..."
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
echo "Step 6/7: Setting up auto-start service..."

# Get current user
CURRENT_USER=$(whoami)
HOME_DIR=$(eval echo ~$CURRENT_USER)

sudo tee /etc/systemd/system/voter-analytics-2-2.service > /dev/null << EOF
[Unit]
Description=Voter Analytics Web UI
After=network.target

[Service]
Type=simple
User=$CURRENT_USER
WorkingDirectory=$HOME_DIR/voter_analytics
Environment="PATH=$HOME_DIR/voter_analytics/venv/bin"
ExecStart=$HOME_DIR/voter_analytics/venv/bin/streamlit run cloud/voter_processor_ui.py --server.port $STREAMLIT_PORT --server.address 0.0.0.0
Restart=always
RestartSec=10

[Install]
WantedBy=multi-user.target
EOF

sudo systemctl daemon-reload
print_status "Auto-start service configured"

# Step 7: Configure firewall (if ufw is available)
echo ""
echo "Step 7/7: Configuring firewall..."
if command -v ufw &> /dev/null; then
    sudo ufw allow ${STREAMLIT_PORT}/tcp 2>/dev/null || true
    print_status "Firewall rule added for port $STREAMLIT_PORT"
else
    print_warning "UFW not installed, skipping firewall configuration"
fi

# Print summary
echo ""
echo "=============================================="
echo -e "${GREEN}  Setup Complete!${NC}"
echo "=============================================="
echo ""
echo "Quick Start (if files already exist):"
echo ""
echo "  cd ~/voter_analytics_2.0"
echo "  source venv/bin/activate"
echo "  streamlit run cloud/voter_processor_ui.py --server.port $STREAMLIT_PORT --server.address 0.0.0.0"
echo ""
echo "OR use systemd service:"
echo ""
echo "  sudo systemctl start voter-analytics-2"
echo "  sudo systemctl enable voter-analytics-2"
echo ""
echo "Access Web UI at: http://YOUR_VM_IP:$STREAMLIT_PORT"
echo ""
echo "----------------------------------------------"
echo "GCP Firewall (if not done):"
echo "  gcloud compute firewall-rules create allow-streamlit \\"
echo "    --allow tcp:$STREAMLIT_PORT --direction INGRESS"
echo ""
echo "----------------------------------------------"
echo "To update code later:"
echo "  cd ~/voter_analytics_2.0 && git pull origin main"
echo ""
echo "=============================================="
echo ""

# Verification
echo "Verifying installation..."
echo ""
echo "Git version:"
git --version
echo ""
echo "Tesseract version:"
tesseract --version | head -1
echo ""
echo "Tesseract languages:"
tesseract --list-langs 2>&1 | grep -E "tam|eng" && print_status "Tamil & English language packs installed"
echo ""
echo "Python version:"
python3 --version
echo ""
echo "Pip packages:"
pip list | grep -E "streamlit|openpyxl|pymupdf|pytesseract|Pillow|requests" || true
echo ""
print_status "Setup verification complete"
