# noScribe Project Quickstart

This file contains the common terminal commands you'll need to run this analytical pipeline.

## 1. Entering the Workspace
Always run this command first when you open a new terminal to ensure your environment is active:
```bash
source venv/bin/activate
```

## 2. Launching the GUI (Recommended)
The easiest way to use the new "Research Edition" tools is through the integrated GUI:
```bash
python3 pipeline_gui.py
```
*Use the sidebar to move through **Download**, **Prepare**, **Transcribe**, **Excel**, and **Graphs**.*

## 3. Background Transcription (CLI)
If you have a very large file and prefer to transcribe via the terminal:
```bash
python3 noScribe.py videos/your_file.mp3 transcripts/output.txt --no-gui
```

## 4. Running Individual Analysis Scripts
If you prefer not to use the GUI, you can run the analytical components manually:

### A. Prepare Audio (Backend)
```bash
python3 analysis/prepare_audio.py input.wav output.mp3 --normalize --trim 0:00 5:00
```

### B. Convert Transcript to Excel (with Lexical Metrics)
```bash
python3 analysis/parse_to_excel.py transcripts/your_output.txt exports/your_analysis.xlsx --lexical --latex
```

### C. Generate Visualization Graphs (with Word Clouds)
```bash
python3 analysis/visualize_data.py exports/your_analysis.xlsx analysis/graphs_folder --all
```

---
**Tip:** You can copy and paste these commands directly into your terminal!
