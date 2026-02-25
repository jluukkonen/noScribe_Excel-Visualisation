#!/usr/bin/env python3
"""
🔧 Audio Preparation Tool
Convert, trim, normalize, and split audio files.
Uses pydub + ffmpeg under the hood.
"""

import os
import sys
import argparse
from pydub import AudioSegment
from pydub.effects import normalize


def time_to_ms(time_str):
    """Convert HH:MM:SS or MM:SS to milliseconds."""
    parts = time_str.strip().split(":")
    if len(parts) == 3:
        h, m, s = int(parts[0]), int(parts[1]), float(parts[2])
        return int((h * 3600 + m * 60 + s) * 1000)
    elif len(parts) == 2:
        m, s = int(parts[0]), float(parts[1])
        return int((m * 60 + s) * 1000)
    else:
        return int(float(parts[0]) * 1000)


def prepare_audio(input_path, output_path, trim_start=None, trim_stop=None,
                  do_normalize=False, target_format="mp3", bitrate="192k"):
    """Process an audio file with optional trimming and normalization."""
    
    ext = os.path.splitext(input_path)[1].lower().lstrip(".")
    format_map = {"m4a": "m4a", "mp4": "mp4", "wav": "wav", "flac": "flac",
                  "ogg": "ogg", "wma": "wma", "aac": "aac", "mp3": "mp3"}
    input_fmt = format_map.get(ext, ext)
    
    print(f"Loading: {input_path}")
    if input_fmt == "m4a":
        audio = AudioSegment.from_file(input_path, format="m4a")
    elif input_fmt == "mp4":
        audio = AudioSegment.from_file(input_path, format="mp4")
    else:
        audio = AudioSegment.from_file(input_path, format=input_fmt)
    
    duration_s = len(audio) / 1000
    print(f"  Duration: {duration_s:.1f}s ({duration_s/60:.1f} min)")
    print(f"  Channels: {audio.channels}, Sample rate: {audio.frame_rate}Hz")
    
    # Trim
    if trim_start or trim_stop:
        start_ms = time_to_ms(trim_start) if trim_start else 0
        stop_ms = time_to_ms(trim_stop) if trim_stop else len(audio)
        print(f"  Trimming: {trim_start or '0:00'} → {trim_stop or 'end'}")
        audio = audio[start_ms:stop_ms]
        print(f"  New duration: {len(audio)/1000:.1f}s")
    
    # Normalize volume
    if do_normalize:
        print("  Normalizing volume...")
        audio = normalize(audio)
    
    # Export
    print(f"  Exporting as {target_format.upper()} → {output_path}")
    os.makedirs(os.path.dirname(output_path) or ".", exist_ok=True)
    
    export_params = {"format": target_format}
    if target_format == "mp3":
        export_params["bitrate"] = bitrate
    
    audio.export(output_path, **export_params)
    
    output_size = os.path.getsize(output_path) / (1024 * 1024)
    print(f"  ✅ Saved: {output_path} ({output_size:.1f} MB)")
    return output_path


def split_on_silence_segments(input_path, output_dir, min_silence_len=1000,
                               silence_thresh=-40, min_chunk_len=30000):
    """Split audio at silence gaps into separate files."""
    from pydub.silence import split_on_silence as pydub_split
    
    ext = os.path.splitext(input_path)[1].lower().lstrip(".")
    audio = AudioSegment.from_file(input_path, format=ext if ext != "m4a" else "m4a")
    
    print(f"Splitting on silence (min gap: {min_silence_len}ms, threshold: {silence_thresh}dB)...")
    chunks = pydub_split(audio, min_silence_len=min_silence_len,
                          silence_thresh=silence_thresh,
                          keep_silence=500)
    
    # Merge chunks that are too short
    merged = []
    current = chunks[0] if chunks else AudioSegment.empty()
    for chunk in chunks[1:]:
        if len(current) < min_chunk_len:
            current += chunk
        else:
            merged.append(current)
            current = chunk
    if len(current) > 0:
        merged.append(current)
    
    os.makedirs(output_dir, exist_ok=True)
    basename = os.path.splitext(os.path.basename(input_path))[0]
    
    paths = []
    for i, chunk in enumerate(merged):
        path = os.path.join(output_dir, f"{basename}_part{i+1:02d}.mp3")
        chunk.export(path, format="mp3", bitrate="192k")
        print(f"  Part {i+1}: {len(chunk)/1000:.1f}s → {path}")
        paths.append(path)
    
    print(f"\n✅ Split into {len(merged)} parts in {output_dir}/")
    return paths


def main():
    parser = argparse.ArgumentParser(description="Prepare audio files for transcription")
    parser.add_argument("input_file", help="Input audio file")
    parser.add_argument("output_file", help="Output audio file path")
    
    parser.add_argument("--format", default="mp3", choices=["mp3", "wav", "flac", "ogg"],
                        help="Output format (default: mp3)")
    parser.add_argument("--start", default=None, help="Trim start (HH:MM:SS)")
    parser.add_argument("--stop", default=None, help="Trim stop (HH:MM:SS)")
    parser.add_argument("--normalize", action="store_true", help="Normalize volume")
    parser.add_argument("--bitrate", default="192k", help="MP3 bitrate (default: 192k)")
    parser.add_argument("--split", action="store_true",
                        help="Split on silence into separate files")
    
    args = parser.parse_args()
    
    if args.split:
        output_dir = os.path.splitext(args.output_file)[0] + "_parts"
        split_on_silence_segments(args.input_file, output_dir)
    else:
        prepare_audio(
            args.input_file, args.output_file,
            trim_start=args.start, trim_stop=args.stop,
            do_normalize=args.normalize,
            target_format=args.format, bitrate=args.bitrate
        )


if __name__ == "__main__":
    main()
