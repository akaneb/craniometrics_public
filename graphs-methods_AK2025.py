import pandas as pd
import matplotlib.pyplot as plt

# File paths
uc_file = r'C:\Users\akane\Desktop\cm_data\avg_UC.xlsx'
plagio_file = r'C:\Users\akane\Desktop\cm_data\avg_plagio.xlsx'

# Highlighted rows (0-based)
highlight_rows = {
    'LR': [33, 68, 75],        # rows 34, 69, 76
    'Height': [67, 11, 47],    # rows 68, 12, 48
    'AntPost': [4, 43, 17],    # rows 5, 44, 18
}

# Sheets to use
sheets = ['LR', 'Height', 'AntPost']

# Create 3 subplots: Axial, Sagittal, Coronal
fig, (ax1, ax2, ax3) = plt.subplots(1, 3, figsize=(20, 8))

for sheet in sheets:
    # Read UC and Plagio data
    uc_df = pd.read_excel(uc_file, sheet_name=sheet).dropna(subset=['x', 'y', 'z'])
    plagio_df = pd.read_excel(plagio_file, sheet_name=sheet).dropna(subset=['x', 'y', 'z'])

    # Axial view (X vs Z)
    ax1.scatter(uc_df['x'], uc_df['z'], c='#FFE066', s=10, alpha=0.4, label='UC Outline' if sheet == 'LR' else "")
    ax1.scatter(plagio_df['x'], plagio_df['z'], c='lightgray', s=10, alpha=0.4, label='Plagio Outline' if sheet == 'LR' else "")

    # Sagittal view (Z vs Y)
    ax2.scatter(uc_df['z'], uc_df['y'], c='#FFE066', s=10, alpha=0.4, label='UC Outline' if sheet == 'LR' else "")
    ax2.scatter(plagio_df['z'], plagio_df['y'], c='lightgray', s=10, alpha=0.4, label='Plagio Outline' if sheet == 'LR' else "")

    # Coronal view (X vs Y)
    ax3.scatter(uc_df['x'], uc_df['y'], c='#FFE066', s=10, alpha=0.4, label='UC Outline' if sheet == 'LR' else "")
    ax3.scatter(plagio_df['x'], plagio_df['y'], c='lightgray', s=10, alpha=0.4, label='Plagio Outline' if sheet == 'LR' else "")

    # Highlighted points
    uc_high = uc_df.iloc[highlight_rows[sheet]]
    plagio_high = plagio_df.iloc[highlight_rows[sheet]]

    # Axial highlights (X vs Z)
    ax1.scatter(uc_high['x'], uc_high['z'], c='#FFD700', s=80, marker='o', label='UC Points' if sheet == 'LR' else "")
    ax1.scatter(plagio_high['x'], plagio_high['z'], c='black', s=80, marker='o', label='Plagio Points' if sheet == 'LR' else "")

    # Sagittal highlights (Z vs Y)
    ax2.scatter(uc_high['z'], uc_high['y'], c='#FFD700', s=80, marker='o', label='UC Points' if sheet == 'LR' else "")
    ax2.scatter(plagio_high['z'], plagio_high['y'], c='black', s=80, marker='o', label='Plagio Points' if sheet == 'LR' else "")

    # Coronal highlights (X vs Y)
    ax3.scatter(uc_high['x'], uc_high['y'], c='#FFD700', s=80, marker='o', label='UC Points' if sheet == 'LR' else "")
    ax3.scatter(plagio_high['x'], plagio_high['y'], c='black', s=80, marker='o', label='Plagio Points' if sheet == 'LR' else "")

# Axial settings
ax1.set_title('Axial View (X vs Z)')
ax1.set_xlabel('X (Width: L ↔ R)')
ax1.set_ylabel('Z (Length: Ant ↔ Post)', labelpad=-10)
ax1.yaxis.set_label_coords(-0.1, 0.5)
ax1.set_aspect('equal')
ax1.grid(True)
ax1.tick_params(labelsize=10)

# Sagittal settings
ax2.set_title('Sagittal View (Z vs Y)')
ax2.set_xlabel('Z (Length: Ant ↔ Post)')
ax2.set_ylabel('Y (Height: Inf ↔ Sup)', labelpad=-10)
ax2.yaxis.set_label_coords(-0.1, 0.5)
ax2.set_aspect('equal')
ax2.grid(True)
ax2.tick_params(labelsize=10)

# Coronal settings
ax3.set_title('Coronal View (X vs Y)')
ax3.set_xlabel('X (Width: L ↔ R)')
ax3.set_ylabel('Y (Height: Inf ↔ Sup)', labelpad=-10)
ax3.yaxis.set_label_coords(-0.1, 0.5)
ax3.set_aspect('equal')
ax3.grid(True)
ax3.tick_params(labelsize=10)

# Leave space at bottom for global legend
plt.tight_layout(rect=[0, 0.1, 1, 1])

# Gather all legend entries
handles1, labels1 = ax1.get_legend_handles_labels()
handles2, labels2 = ax2.get_legend_handles_labels()
handles3, labels3 = ax3.get_legend_handles_labels()

# Combine and deduplicate
all_handles = handles1 + handles2 + handles3
all_labels = labels1 + labels2 + labels3
unique = dict(zip(all_labels, all_handles))

# Global legend at bottom center
fig.legend(
    unique.values(),
    unique.keys(),
    loc='lower center',
    ncol=4,
    bbox_to_anchor=(0.5, 0.02),
    frameon=False
)

plt.show()
