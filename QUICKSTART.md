# noScribe Project Quickstart

This file contains the common terminal commands you'll need to run this project.

## 1. Entering the Workspace
Always run this command first when you open a new terminal:
```bash
cd /Users/joonasluukkonen/Desktop/NoScribe
source venv/bin/activate
```

## 2. Launching the GUI
If you want to use the graphical window:
```bash
python3 noScribe.py
```

## 3. Background Transcription (CLI)
To transcribe a file without the GUI (faster/efficient for long files):
```bash
python3 noScribe.py videos/your_file.mp3 transcripts/output_name.txt --no-gui
```

## 4. Downloading Audio from YouTube
```bash
./download_audio.sh "https://www.youtube.com/watch?v=..."
```
*(The audio will be saved in the `videos/` folder)*

## 5. Running the Analysis Pipeline
Once you have a transcript in the `transcripts/` folder:

### A. Convert Transcript to Excel
```bash
python3 analysis/parse_to_excel.py transcripts/your_transcript.txt exports/your_analysis.xlsx
```

### B. Generate Visualization Graphs
```bash
python3 analysis/visualize_data.py exports/your_analysis.xlsx analysis/graphs_folder
```

---
**Tip:** You can just copy and paste these commands directly into your terminal!
