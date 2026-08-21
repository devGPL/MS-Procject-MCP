#!/usr/bin/env bash
# Run both refactor gates plus a compile check. Intended to run after every
# commit of the server.py split, on macOS/Linux, with no MS Project present.
#
#     ./tools/gates.sh
#
# Exit 0 only if all three pass.
set -uo pipefail

cd "$(dirname "$0")/.." || exit 2

fail=0

echo "== Gate 1: unresolved names =="
python3 tools/check_names.py ./*.py || fail=1

echo
echo "== Gate 2: tool registration vs baseline =="
python3 tools/snap_tools.py . /tmp/snap_tools_current.json \
    --baseline tools/baseline_tools.json || fail=1

echo
echo "== Gate 3: dead imports and orphan banners =="
python3 tools/check_unused.py ./*.py || fail=1

echo
echo "== compileall =="
if python3 -m compileall -q ./*.py; then
    echo "OK"
else
    fail=1
fi

echo
if [ "$fail" -eq 0 ]; then
    echo "GATES: PASS"
else
    echo "GATES: FAIL"
fi
exit "$fail"
