#!/bin/sh
set -eu

APP_DIR="/app/share/dataflow"

PYTHON_VERSION="$(python3 -c 'import sys; print("{}.{}".format(sys.version_info.major, sys.version_info.minor))')"
SITE_PACKAGES="/app/lib/python${PYTHON_VERSION}/site-packages"

if [ -d "$SITE_PACKAGES" ]; then
    export PYTHONPATH="${SITE_PACKAGES}${PYTHONPATH:+:$PYTHONPATH}"
fi

cd "$APP_DIR"
exec python3 dataflow.py "$@"
