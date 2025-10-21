#!/bin/bash

echo "==============================="
echo "    mykLabs Streamlit Application Launcher"
echo "==============================="

# Check if virtual environment exists
if [ ! -f ".venv/bin/activate" ]; then
    echo "Error: Virtual environment .venv not found"
    echo "Please create virtual environment first: python -m venv .venv"
    exit 1
fi

# Activate virtual environment
echo "Activating virtual environment..."
source .venv/bin/activate

# Check if streamlit is installed
if ! python -c "import streamlit" &> /dev/null; then
    echo "Error: streamlit not installed in virtual environment"
    echo "Please install: pip install streamlit"
    exit 1
fi

# Check if webui.py exists
if [ ! -f "./webui/webui.py" ]; then
    echo "Error: webui.py file not found"
    echo "Please ensure script is running in the correct directory"
    exit 1
fi

echo "Virtual environment activated successfully!"
echo "Starting Streamlit application..."
echo "Application will open automatically in your browser"
echo "Press Ctrl+C to stop the service"

# Run the streamlit application
python3 -m streamlit run webui/webui.py