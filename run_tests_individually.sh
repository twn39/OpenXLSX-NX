#!/bin/bash
cd build
./OpenXLSXTests --list-tests | grep -E '^  [a-zA-Z0-9]' | sed -e 's/^[[:space:]]*//' | while read test; do
    echo "Running $test"
    OUTPUT=$(./OpenXLSXTests "$test" 2>&1)
    res=$?
    if [ $res -ne 0 ]; then
        echo "FAILED: $test with exit code $res"
        echo "OUTPUT:"
        echo "$OUTPUT"
        echo "----------------------------------------"
    fi
done
