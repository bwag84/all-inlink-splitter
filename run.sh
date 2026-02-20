#!/usr/bin/env bash
#
# Convenience wrapper — builds the Docker image and runs the splitter.
# All CLI arguments are forwarded to splitter.py.
#
# Usage:
#   ./run.sh                        # default region split
#   ./run.sh --split url            # URL-pattern split
#   ./run.sh --split both --no-filter
#

set -euo pipefail

IMAGE_NAME="all-inlink-splitter"
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"

docker build -t "$IMAGE_NAME" "$SCRIPT_DIR"

docker run --rm \
  -v "$SCRIPT_DIR/input:/app/input" \
  -v "$SCRIPT_DIR/output:/app/output" \
  "$IMAGE_NAME" "$@"
