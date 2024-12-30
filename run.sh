#!/bin/bash

set -eu

echo
echo Merging with arguments: "$@"
build/install/poi-mail-merge/bin/poi-mail-merge "$@"
