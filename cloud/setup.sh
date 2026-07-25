#!/bin/bash
#
# Voter Analytics - One-Time Setup Script
# Run this on a fresh Ubuntu 22.04 VM
#
# Usage:
#   # Default port 8501:
#   curl -sSL https://raw.githubusercontent.com/thiraviyalingesh/voters_project_2.0/main/cloud/setup.sh | tr -d '\r' | bash
#
#   # Custom port (e.g., 8080):
#   curl -sSL https://raw.githubusercontent.com/thiraviyalingesh/voters_project_2.0/main/cloud/setup.sh | tr -d '\r' | bash -s -- --port 8080
#

set -e  # Exit on error

# ============== CONFIGURATION ==============
# Change this to use a different port
STREAMLIT_PORT="${STREAMLIT_PORT:-8501}"

# Install location. Must match BASE_DIR in cloud/voter_processor_ui.py
APP_DIR="$HOME/voter_analytics_2.0"

# systemd unit name (used as $SERVICE_NAME.service)
SERVICE_NAME="voter-analytics-2"

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
echo "Install directory: $APP_DIR"
if [ -d "$APP_DIR" ]; then
    print_warning "Directory exists. Pulling latest..."
    cd "$APP_DIR" && git pull origin main || true
else
    git clone "$REPO_URL" "$APP_DIR"
fi

# Create required directories
mkdir -p "$APP_DIR/uploads"
mkdir -p "$APP_DIR/uploads/output"
mkdir -p "$APP_DIR/processing"
mkdir -p "$APP_DIR/output"
mkdir -p "$APP_DIR/logs"
cd "$APP_DIR"
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
    watchdog \
    pandas \
    matplotlib

print_status "Python packages installed"

# Step 6: Create systemd service for auto-start
echo ""
echo "Step 6/7: Setting up auto-start service..."

# Get current user
CURRENT_USER=$(whoami)

# Remove the old, misnamed unit from previous versions of this script
if [ -f /etc/systemd/system/voter-analytics-2-2.service ]; then
    sudo systemctl disable --now voter-analytics-2-2.service 2>/dev/null || true
    sudo rm -f /etc/systemd/system/voter-analytics-2-2.service
    print_warning "Removed stale voter-analytics-2-2.service"
fi

# WorkingDirectory is cloud/ so Streamlit picks up cloud/.streamlit/config.toml
# (maxUploadSize = 5000 MB). PATH must include the system dirs so pytesseract
# can find the tesseract binary.
sudo tee /etc/systemd/system/${SERVICE_NAME}.service > /dev/null << EOF
[Unit]
Description=Voter Analytics Web UI
After=network.target

[Service]
Type=simple
User=$CURRENT_USER
WorkingDirectory=$APP_DIR/cloud
Environment="PATH=$APP_DIR/venv/bin:/usr/local/sbin:/usr/local/bin:/usr/sbin:/usr/bin:/sbin:/bin"
ExecStart=$APP_DIR/venv/bin/streamlit run voter_processor_ui.py --server.port $STREAMLIT_PORT --server.address 0.0.0.0
Restart=always
RestartSec=10

[Install]
WantedBy=multi-user.target
EOF

sudo systemctl daemon-reload
print_status "Auto-start service configured ($SERVICE_NAME.service)"

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
echo "Quick Start (run from cloud/ so .streamlit/config.toml applies):"
echo ""
echo "  cd $APP_DIR/cloud"
echo "  source ../venv/bin/activate"
echo "  streamlit run voter_processor_ui.py --server.port $STREAMLIT_PORT --server.address 0.0.0.0"
echo ""
echo "OR use systemd service (recommended, auto-restarts):"
echo ""
echo "  sudo systemctl enable --now $SERVICE_NAME"
echo "  sudo systemctl status $SERVICE_NAME"
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
echo "  cd $APP_DIR && git pull origin main && sudo systemctl restart $SERVICE_NAME"
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
