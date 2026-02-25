#!/bin/bash

# A simple wrapper to download YouTube audio directly into the videos/ folder for noScribe

if [ "$#" -ne 1 ]; then
    echo "Usage: ./download_audio.sh <youtube_url>"
    exit 1
fi

URL=$1
OUTPUT_DIR="videos"

echo "Fetching audio from: $URL"
echo "Saving to: $OUTPUT_DIR/"

# Use yt-dlp to download the best audio format and convert to mp3
yt-dlp -x --audio-format mp3 \
    -o "$OUTPUT_DIR/%(title)s.%(ext)s" \
    --restrict-filenames \
    --force-overwrites \
    "$URL"

echo ""
echo "Done! The audio file is now in the '$OUTPUT_DIR/' folder ready for noScribe."
