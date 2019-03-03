#!/bin/sh

set -e
set -o pipefail

echo "{\"scriptId\":\"$1\"}" > .clasp.json

clasp push

