#!/bin/bash
# Build Windows EXE via WSL interop
# Automatically converts between Linux and Windows paths

set -e

echo "============================================"
echo "  WSL Windows Build Launcher"
echo "============================================"
echo ""

# Get script directory (Linux path)
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"

# Convert to Windows path format for cmd
WIN_SCRIPT_DIR=$(wslpath -w "$SCRIPT_DIR")
WIN_BAT="${WIN_SCRIPT_DIR}\\build_windows.bat"

echo "Script directory (Linux): $SCRIPT_DIR"
echo "Script directory (Windows): $WIN_SCRIPT_DIR"
echo "Executing: $WIN_BAT"
echo ""

# Check if Windows cmd.exe is available
if [ ! -f "/mnt/c/Windows/System32/cmd.exe" ]; then
    echo "[ERROR] Windows cmd.exe not found at expected location"
    echo "        WSL interop may not be enabled"
    exit 1
fi

# Check if batch file exists
if [ ! -f "$SCRIPT_DIR/build_windows.bat" ]; then
    echo "[ERROR] build_windows.bat not found in: $SCRIPT_DIR"
    exit 1
fi

# Run the batch file via Windows cmd
/mnt/c/Windows/System32/cmd.exe /c "$WIN_BAT"

if [ $? -ne 0 ]; then
    echo ""
    echo "[ERROR] Build failed"
    exit 1
fi

echo ""
EXE_PATH="$SCRIPT_DIR/dist/cpp_highlight.exe"
if [ -f "$EXE_PATH" ]; then
    echo "[SUCCESS] Build completed:"
    ls -lh "$EXE_PATH"
else
    echo "[WARNING] Expected output not found: $EXE_PATH"
fi
