#!/bin/bash

SOURCE="../path-source-location"
DEST="/path-target-location"

mkdir -p "$DEST"

while IFS= read -r file; do
    file="${file%$'\r'}"
    cp "$SOURCE/$file" "$DEST/"
done < bangle-mxd.txt

echo "Selesai."