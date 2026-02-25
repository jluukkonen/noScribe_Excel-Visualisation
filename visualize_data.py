import pandas as pd
import matplotlib.pyplot as plt
import seaborn as plt_sns
import os
import sys

def create_visualizations(excel_path, output_dir):
    print(f"Reading data from {excel_path}...")
    
    # Read the data sheet
    df = pd.read_excel(excel_path, sheet_name="Transcript Data")
    
    # Add a Turn Number column for the X-axis
    df['Turn'] = range(1, len(df) + 1)
    
    # Dynamically assign colors to all unique speakers
    speakers = list(df['Speaker'].unique())
    colors = plt_sns.color_palette("husl", len(speakers)).as_hex()
    color_palette = dict(zip(speakers, colors))
    
    # Ensure output directory exists
    os.makedirs(output_dir, exist_ok=True)
    
    # Dynamic scaling for large datasets
    num_turns = len(df)
    plot_width = max(14, num_turns / 8) # Stretch the plot if there are many turns
    x_tick_step = max(5, num_turns // 20) # Avoid overlapping x labels
    
    # ==========================================
    # 1. The Conversational Timeline (Word Count per Turn)
    # ==========================================
    plt.figure(figsize=(plot_width, 6))
    
    # Create a bar plot showing the length of each turn
    plt_sns.barplot(data=df, x='Turn', y='Word Count', hue='Speaker', palette=color_palette, dodge=False)
    
    plt.title('Conversational Rhythm: Word Count per Turn', fontsize=16, pad=20)
    plt.xlabel('Speaking Turn (Chronological ->)', fontsize=12)
    plt.ylabel('Words Spoken (Length of Turn)', fontsize=12)
    
    # Clean up the X-axis labels based on dynamic step
    plt.xticks(range(0, num_turns, x_tick_step))
    
    plt.grid(axis='y', linestyle='--', alpha=0.7)
    plt.legend(title='Speaker')
    
    timeline_path = os.path.join(output_dir, "conversational_timeline.png")
    plt.tight_layout()
    plt.savefig(timeline_path, dpi=300)
    print(f"Saved: {timeline_path} (Width: {plot_width})")
    plt.close()
    
    # ==========================================
    # 2. Friction & Hesitation Scatterplot
    # ==========================================
    plt.figure(figsize=(plot_width, 6))
    
    # We want to plot dots where disfluencies or long pauses happened
    df_friction = df[(df['Disfluencies (uh/um/yeah...)'] > 0) | (df['Long Pauses'] > 0)]
    
    # Dynamically scale the size of the dots so they don't overlap in dense graphs
    dot_scale = 100 if num_turns < 50 else 30
    
    # Plot Disfluencies
    plt.scatter(df_friction['Turn'], df_friction['Disfluencies (uh/um/yeah...)'], 
                color='#E74C3C', s=df_friction['Disfluencies (uh/um/yeah...)']*dot_scale, 
                alpha=0.6, label='Disfluencies (uh/um)')
                
    # Plot Long Pauses
    plt.scatter(df_friction['Turn'], df_friction['Long Pauses'], 
                color='#F39C12', s=df_friction['Long Pauses']*dot_scale, 
                alpha=0.8, marker='s', label='Long Pauses (..)')
    
    # Add a faint background line showing who was speaking
    bg_alpha = 0.1 if num_turns < 50 else 0.03 # Make the background lines much fainter on dense plots
    for i, row in df.iterrows():
        color = color_palette.get(row['Speaker'], "#CCCCCC")
        plt.axvspan(row['Turn']-0.5, row['Turn']+0.5, color=color, alpha=bg_alpha)
        
    plt.title('Conversational Friction: Where Hesitations and Pauses Occurred', fontsize=16, pad=20)
    plt.xlabel('Speaking Turn', fontsize=12)
    plt.ylabel('Count in that Turn', fontsize=12)
    
    plt.xticks(range(0, num_turns, x_tick_step))
        
    plt.legend()
    plt.grid(True, linestyle=':', alpha=0.6)
    
    friction_path = os.path.join(output_dir, "friction_heatmap.png")
    plt.tight_layout()
    plt.savefig(friction_path, dpi=300)
    print(f"Saved: {friction_path} (Width: {plot_width})")
    plt.close()
    
def main():
    if len(sys.argv) < 3:
        print("Usage: python visualize_data.py <input_excel> <output_dir>")
        return
        
    excel_file = sys.argv[1]
    out_dir = sys.argv[2]
    
    create_visualizations(excel_file, out_dir)

if __name__ == "__main__":
    main()
