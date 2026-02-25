import pandas as pd
import matplotlib.pyplot as plt
import seaborn as plt_sns
import os
import sys
import argparse

def create_visualizations(excel_path, output_dir, charts=None, fmt="png", dpi=300):
    """Generate selected visualizations from transcript Excel data.
    
    Args:
        charts: list of chart names to generate. If None, generates all.
                 Options: timeline, friction, pie, distribution
        fmt: output format (png, svg, pdf)
        dpi: output resolution
    """
    print(f"Reading data from {excel_path}...")
    
    df = pd.read_excel(excel_path, sheet_name="Transcript Data")
    df['Turn'] = range(1, len(df) + 1)
    
    speakers = list(df['Speaker'].unique())
    colors = plt_sns.color_palette("husl", len(speakers)).as_hex()
    color_palette = dict(zip(speakers, colors))
    
    os.makedirs(output_dir, exist_ok=True)
    
    num_turns = len(df)
    plot_width = max(14, num_turns / 8)
    x_tick_step = max(5, num_turns // 20)
    
    all_charts = {"timeline", "friction", "pie", "distribution", "matrix", "wordcloud"}
    if charts is None:
        charts = all_charts
    else:
        charts = set(charts) & all_charts

    generated = []

    # ==========================================
    # 1. Conversational Timeline (Word Count per Turn)
    # ==========================================
    if "timeline" in charts:
        plt.figure(figsize=(plot_width, 6))
        plt_sns.barplot(data=df, x='Turn', y='Word Count', hue='Speaker',
                        palette=color_palette, dodge=False)
        plt.title('Conversational Rhythm: Word Count per Turn', fontsize=16, pad=20)
        plt.xlabel('Speaking Turn (Chronological ->)', fontsize=12)
        plt.ylabel('Words Spoken (Length of Turn)', fontsize=12)
        plt.xticks(range(0, num_turns, x_tick_step))
        plt.grid(axis='y', linestyle='--', alpha=0.7)
        plt.legend(title='Speaker')
        path = os.path.join(output_dir, f"conversational_timeline.{fmt}")
        plt.tight_layout()
        plt.savefig(path, dpi=dpi)
        print(f"  ✅ {path}")
        plt.close()
        generated.append(path)
    
    # ==========================================
    # 2. Friction & Hesitation Scatterplot
    # ==========================================
    if "friction" in charts:
        plt.figure(figsize=(plot_width, 6))
        df_friction = df[(df['Disfluencies (uh/um/yeah...)'] > 0) | (df['Long Pauses'] > 0)]
        dot_scale = 100 if num_turns < 50 else 30
        
        plt.scatter(df_friction['Turn'], df_friction['Disfluencies (uh/um/yeah...)'], 
                    color='#E74C3C', s=df_friction['Disfluencies (uh/um/yeah...)']*dot_scale, 
                    alpha=0.6, label='Disfluencies (uh/um)')
        plt.scatter(df_friction['Turn'], df_friction['Long Pauses'], 
                    color='#F39C12', s=df_friction['Long Pauses']*dot_scale, 
                    alpha=0.8, marker='s', label='Long Pauses (..)')
        
        bg_alpha = 0.1 if num_turns < 50 else 0.03
        for i, row in df.iterrows():
            color = color_palette.get(row['Speaker'], "#CCCCCC")
            plt.axvspan(row['Turn']-0.5, row['Turn']+0.5, color=color, alpha=bg_alpha)
            
        plt.title('Conversational Friction: Where Hesitations and Pauses Occurred', fontsize=16, pad=20)
        plt.xlabel('Speaking Turn', fontsize=12)
        plt.ylabel('Count in that Turn', fontsize=12)
        plt.xticks(range(0, num_turns, x_tick_step))
        plt.legend()
        plt.grid(True, linestyle=':', alpha=0.6)
        path = os.path.join(output_dir, f"friction_heatmap.{fmt}")
        plt.tight_layout()
        plt.savefig(path, dpi=dpi)
        print(f"  ✅ {path}")
        plt.close()
        generated.append(path)

    # ==========================================
    # 3. Speaker Balance Pie Chart
    # ==========================================
    if "pie" in charts:
        fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(14, 6))
        
        # Pie by total words
        word_counts = df.groupby('Speaker')['Word Count'].sum()
        pie_colors = [color_palette.get(s, "#CCCCCC") for s in word_counts.index]
        ax1.pie(word_counts, labels=word_counts.index, autopct='%1.1f%%',
                colors=pie_colors, startangle=90, textprops={'fontsize': 12})
        ax1.set_title('Share of Total Words', fontsize=14, pad=15)
        
        # Pie by total turns
        turn_counts = df['Speaker'].value_counts()
        pie_colors2 = [color_palette.get(s, "#CCCCCC") for s in turn_counts.index]
        ax2.pie(turn_counts, labels=turn_counts.index, autopct='%1.1f%%',
                colors=pie_colors2, startangle=90, textprops={'fontsize': 12})
        ax2.set_title('Share of Speaking Turns', fontsize=14, pad=15)
        
        fig.suptitle('Speaker Balance', fontsize=16, y=1.02)
        path = os.path.join(output_dir, f"speaker_balance.{fmt}")
        plt.tight_layout()
        plt.savefig(path, dpi=dpi, bbox_inches='tight')
        print(f"  ✅ {path}")
        plt.close()
        generated.append(path)

    # ==========================================
    # 4. Turn Length Distribution (Histogram)
    # ==========================================
    if "distribution" in charts:
        fig, ax = plt.subplots(figsize=(10, 6))
        
        for speaker in speakers:
            speaker_words = df[df['Speaker'] == speaker]['Word Count']
            ax.hist(speaker_words, bins=20, alpha=0.6, label=speaker,
                    color=color_palette.get(speaker, "#CCCCCC"), edgecolor='white')
        
        ax.set_title('Turn Length Distribution: How Long Are Speaking Turns?', fontsize=16, pad=20)
        ax.set_xlabel('Words per Turn', fontsize=12)
        ax.set_ylabel('Frequency', fontsize=12)
        ax.legend(title='Speaker')
        ax.grid(axis='y', linestyle='--', alpha=0.5)
        path = os.path.join(output_dir, f"turn_distribution.{fmt}")
        plt.tight_layout()
        plt.savefig(path, dpi=dpi)
        print(f"  ✅ {path}")
        plt.close()
        generated.append(path)

    # ==========================================
    # 5. Turn-Taking Adjacency Matrix
    # ==========================================
    if "matrix" in charts and len(speakers) >= 2:
        import numpy as np

        # Build transition matrix: who speaks after whom
        transitions = {s: {t: 0 for t in speakers} for s in speakers}
        speaker_list = df['Speaker'].tolist()
        for i in range(len(speaker_list) - 1):
            current = speaker_list[i]
            next_spk = speaker_list[i + 1]
            if current != next_spk:  # only count actual speaker changes
                transitions[current][next_spk] += 1

        # Convert to percentage (row-normalized)
        matrix_data = []
        labels = speakers
        for s in labels:
            row_total = sum(transitions[s].values())
            if row_total > 0:
                matrix_data.append([transitions[s][t] / row_total * 100 for t in labels])
            else:
                matrix_data.append([0] * len(labels))

        matrix_arr = np.array(matrix_data)

        fig, ax = plt.subplots(figsize=(max(8, len(speakers) * 2.5), max(6, len(speakers) * 2)))

        im = ax.imshow(matrix_arr, cmap='YlOrRd', aspect='auto', vmin=0)
        fig.colorbar(im, ax=ax, label='Transition %', shrink=0.8)

        ax.set_xticks(range(len(labels)))
        ax.set_xticklabels(labels, fontsize=12, rotation=45, ha='right')
        ax.set_yticks(range(len(labels)))
        ax.set_yticklabels(labels, fontsize=12)

        ax.set_xlabel('Next Speaker  →', fontsize=12)
        ax.set_ylabel('Current Speaker  →', fontsize=12)
        ax.set_title('Turn-Taking Adjacency Matrix:\nWho Responds to Whom?', fontsize=16, pad=20)

        # Annotate cells with percentages
        for i in range(len(labels)):
            for j in range(len(labels)):
                val = matrix_arr[i, j]
                if val > 0:
                    text_color = 'white' if val > 50 else 'black'
                    ax.text(j, i, f'{val:.0f}%', ha='center', va='center',
                            fontsize=11, fontweight='bold', color=text_color)

        path = os.path.join(output_dir, f"turn_taking_matrix.{fmt}")
        plt.tight_layout()
        plt.savefig(path, dpi=dpi, bbox_inches='tight')
        print(f"  ✅ {path}")
        plt.close()
        generated.append(path)

    # ==========================================
    # 6. Word Clouds per Speaker
    # ==========================================
    if "wordcloud" in charts:
        from wordcloud import WordCloud
        import string

        stop_words = {"the","a","an","and","or","but","in","on","at","to","for","of","is","it","i","you","he","she",
                      "we","they","that","this","was","were","are","be","been","being","have","has","had","do","does",
                      "did","will","would","could","should","may","might","can","shall","so","if","then","than",
                      "not","no","my","your","our","their","its","me","him","her","us","them","what","which","who",
                      "where","when","how","just","also","very","with","from","about","as","all","up","out",
                      "uh","um","hmm","yeah","like"}

        for speaker in speakers:
            text_data = " ".join(df[df["Speaker"] == speaker]["Text"].tolist()).lower()
            
            # Clean text
            for p in string.punctuation:
                if p not in ("'", "-"):
                    text_data = text_data.replace(p, " ")
            text_data = text_data.replace("(.)", "").replace("(..)", "").replace("(...)", "")
            
            # Only generate if there's enough text
            if len(text_data.split()) < 5:
                continue

            # Base color for this speaker
            base_hex = color_palette.get(speaker, "#333333")
            
            def color_func(word, font_size, position, orientation, random_state=None, **kwargs):
                return base_hex

            wc = WordCloud(width=800, height=400, background_color="white",
                           stopwords=stop_words, color_func=color_func,
                           max_words=100, min_word_length=3).generate(text_data)

            plt.figure(figsize=(10, 5))
            plt.imshow(wc, interpolation="bilinear")
            plt.axis("off")
            plt.title(f"Word Cloud: {speaker}", fontsize=18, pad=20, color=base_hex)
            
            # Clean filename
            safe_speaker = "".join(c for c in speaker if c.isalnum() or c in " -_").strip()
            path = os.path.join(output_dir, f"wordcloud_{safe_speaker}.{fmt}")
            plt.tight_layout()
            plt.savefig(path, dpi=dpi, bbox_inches='tight')
            print(f"  ✅ {path}")
            plt.close()
            generated.append(path)

    print(f"\nGenerated {len(generated)} chart(s) in {output_dir}/")
    return generated


def main():
    parser = argparse.ArgumentParser(description="Generate visualizations from transcript Excel data")
    parser.add_argument("excel_file", help="Input .xlsx file")
    parser.add_argument("output_dir", help="Output directory for charts")
    
    # Chart selection flags
    parser.add_argument("--timeline", action="store_true", help="Conversational timeline bar chart")
    parser.add_argument("--friction", action="store_true", help="Friction & hesitation scatterplot")
    parser.add_argument("--pie", action="store_true", help="Speaker balance pie chart")
    parser.add_argument("--distribution", action="store_true", help="Turn length distribution histogram")
    parser.add_argument("--matrix", action="store_true", help="Turn-taking adjacency matrix heatmap")
    parser.add_argument("--wordcloud", action="store_true", help="Word cloud per speaker")
    parser.add_argument("--all", action="store_true", default=False, help="Generate all charts")
    
    # Output settings
    parser.add_argument("--format", default="png", choices=["png", "svg", "pdf"],
                        help="Output format (default: png)")
    parser.add_argument("--dpi", type=int, default=300, choices=[150, 300, 600],
                        help="Output DPI (default: 300)")

    args = parser.parse_args()

    # Determine which charts to generate
    selected = []
    if args.timeline: selected.append("timeline")
    if args.friction: selected.append("friction")
    if args.pie: selected.append("pie")
    if args.distribution: selected.append("distribution")
    if args.matrix: selected.append("matrix")
    if args.wordcloud: selected.append("wordcloud")
    
    # If none specified or --all, generate all
    if not selected or args.all:
        selected = None  # None = all charts
    
    create_visualizations(args.excel_file, args.output_dir,
                         charts=selected, fmt=args.format, dpi=args.dpi)


if __name__ == "__main__":
    main()
