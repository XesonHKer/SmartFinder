#!/bin/bash
DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
export GS_LIB="$DIR/Resource/Init:$DIR/Resource/lib"
exec "$DIR/gs" "$@"
