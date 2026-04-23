#!/bin/bash
# Bus Reports Portal — auto-start script
echo "🚌 Bus Reports Portal"
echo "Checking dependencies..."

if [ ! -d "node_modules" ]; then
  echo "Installing Node.js dependencies (first run only)..."
  npm install
fi

echo ""
echo "✅ Starting portal at http://localhost:3000"
echo "   Press Ctrl+C to stop"
echo ""
node server.js
