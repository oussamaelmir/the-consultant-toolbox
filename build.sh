#!/bin/bash

chmod +x build.sh

# Exit on first error
set -e

echo "🔧 Installing dependencies..."
npm install --omit=dev

echo "🏗️  Building project for production..."
npm run build

echo "✅ Build complete. Files ready in dist/"
