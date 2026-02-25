# noScribe Analysis Pipeline: Research Edition

This repository contains an expanded analytical toolkit for the `noScribe` transcription engine. These tools transform raw time-aligned transcripts into structured research databases, quantitative metrics, and publication-ready visualizations.

## 🚀 Key Features

### 1. 🔧 Integrated Audio Preparation
A dedicated module to source and prepare high-quality audio before transcription.
* **YouTube Downloader**: Easily pull audio from YouTube links.
* **Format Standardization**: Convert WAV, FLAC, M4A, and MP4 to optimized MP3.
* **Volume Normalization**: Ensure consistent playback loudness across different environments.
* **Silence Splitting**: Automatically detect and split recordings at silence gaps.

### 2. 📊 Advanced Lexical Analysis
Moving beyond basic token counts, this edition computes complex linguistic metrics per speaker:
* **Lexical Diversity**: Type-Token Ratio (TTR) and Measure of Textual Lexical Diversity (MTLD).
* **Stylometrics**: Flesch-Kincaid Readability scores to assess discourse complexity.
* **Conversational Friction**: Tracking disfluencies (uh/um), overlap frequency, and pause patterns.

### 3. 📉 High-Granularity Visualizations
Turn your data into insights with presentation-ready charts:
* **Conversational Timeline**: Word count and turn rhythm mapping.
* **Interaction Heatmaps**: Visual mapping of disfluencies and silences over time.
* **Speaker Balance**: Dynamic pie charts of words and turns.
* **Turn-Taking Matrices**: Visualizing "who responds to whom" (conversational topology).
* **Word Clouds**: Automated vocabulary visualization per speaker.

### 4. 📄 Academic Workflow (LaTeX Export)
Export your findings directly into LaTeX documents with a single click.
* **Pre-formatted Tables**: Uses the `booktabs` package for professional tables.
* **Figure Stubs**: Ready-to-use LaTeX code for all generated plots and word clouds.

---

## 🛠️ Installation & Setup

1. **Install noScribe**: Ensure you have the base [noScribe](https://github.com/kaixxx/noScribe) software installed.
2. **Dependencies**:
   ```bash
   pip install pandas openpyxl matplotlib seaborn lexicalrichness textstat wordcloud pydub
   ```
3. **External Tools**: Ensure `ffmpeg` is installed on your system (required for audio processing).

---

## 🎮 How to Use

### Integrated GUI (Recommended)
Launch the new research pipeline interface:
```bash
python3 pipeline_gui.py
```
This integrated window allows you to move through every step:
1. **Download/Prepare Audio** (Source your files)
2. **Transcribe** (Generate the raw text)
3. **Excel + Theme** (Generate metrics and database)
4. **Graphs** (Generate visualizations and word clouds)

### Command Line Usage
For batch processing, the scripts can be run individually:
* **Excel**: `python3 analysis/parse_to_excel.py input.txt output.xlsx --lexical --latex`
* **Graphs**: `python3 analysis/visualize_data.py input.xlsx output_folder --all`

---
*Developed for researchers and students focused on quantitative linguistics and conversation analysis.*
