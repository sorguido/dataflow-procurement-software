#!/bin/sh
set -eu

APP_DIR="/app/share/dataflow"

export LD_LIBRARY_PATH="/app/lib${LD_LIBRARY_PATH:+:$LD_LIBRARY_PATH}"

if [ -d /app/lib/tcl8.6 ]; then
    export TCL_LIBRARY="/app/lib/tcl8.6"
fi

if [ -d /app/lib/tk8.6 ]; then
    export TK_LIBRARY="/app/lib/tk8.6"
fi

PYTHON_VERSION="$(python3 -c 'import sys; print("{}.{}".format(sys.version_info.major, sys.version_info.minor))')"
SITE_PACKAGES="/app/lib/python${PYTHON_VERSION}/site-packages"

if [ -d "$SITE_PACKAGES" ]; then
    export PYTHONPATH="${SITE_PACKAGES}${PYTHONPATH:+:$PYTHONPATH}"
fi

cd "$APP_DIR"
exec python3 dataflow.py "$@"
