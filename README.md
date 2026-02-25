# noScribe Analysis Tools: Excel & Visualizations

This repository contains custom Python scripts designed to take raw `noScribe` AI transcripts and automatically generate structured Excel databases and high-resolution visual graphs.

These tools are built to help researchers and students instantly move from qualitative text data to quantitative visual analysis of conversational dynamics (like interruptions, hesitations, and dominate speaking time).

## What's Included?
1. **`parse_to_excel.py`:** Reads a raw `.txt` transcript and generates a color-coded, heavily formatted Excel file containing a turn-by-turn breakdown and a mathematical **Results Dashboard**.
2. **`visualize_data.py`:** Ingests the output from the Excel file to generate two presentation-ready graphs:
   * **The Conversational Timeline:** A rhythmic bar chart of word counts per turn.
   * **The Friction Heatmap:** A visual scatterplot mapping the location and intensity of conversational disfluencies and silences.
3. **`download_audio.sh`:** A quick shell script wrapper for `yt-dlp` to easily pull pristine `.mp3` files down from YouTube links for transcription.

## Prerequisites
To use these tools, you need the base transcription software installed.
1. Download and install [noScribe](https://github.com/kaixxx/noScribe) locally.
2. Ensure you have activated your `noScribe` python virtual environment.
3. Install the data science libraries we need for the visualization scripts:
   ```bash
   pip install pandas openpyxl matplotlib seaborn
   ```

## Quick Start Guide

Drop these three files into your main `noScribe` folder alongside `noScribe.py`.

### 1. Download Audio (Optional)
If you have a YouTube link you want to analyze, use the shell script:
```bash
./download_audio.sh "https://www.youtube.com/watch?v=YOUR_LINK"
```

### 2. Transcribe Audio
Use `noScribe` via the terminal to transcribe the file in the background (preventing the GUI from crashing on massive files):
```bash
python3 noScribe.py videos/your_file.mp3 transcripts/your_output.txt --no-gui
```

### 3. Generate the Excel Database
Convert the raw AI text into the color-coded Excel ledger:
```bash
python3 parse_to_excel.py transcripts/your_output.txt exports/your_analysis.xlsx
```

### 4. Generate the Visualization Graphs
Draw the Timeline and Heatmap using the data from the Excel file:
```bash
python3 visualize_data.py exports/your_analysis.xlsx analysis/graphs_folder
```

## Reading the Output
- **Excel:** Open the `.xlsx` file. The first tab is a mathematical summary of who dominated the conversation, how many times they paused, and how many times they interrupted. The second tab is the full color-coded transcript.
- **Timeline Graph:** Taller bars mean more words were spoken before the other person interrupted. Colors dictate the speaker.
- **Heatmap Graph:** Red circles are filler language (uh/um). Yellow squares are intense silences (2+ seconds). The higher up and larger the shape, the more times it happened rapidly in a single turn.

---
*Built for educational and research analysis of high-stakes conversational audio.*
