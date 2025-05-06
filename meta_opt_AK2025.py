import os
import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
from scipy.optimize import minimize
from sklearn.metrics import roc_curve
from scipy import stats
import random
import time

# Directories containing the xlsx files - edit appropriately
UC_dir = r"***"
plagio_dir = r"***"

# Set a random seed using the current time
random_seed = int(time.time())
random.seed(random_seed)
np.random.seed(random_seed)

print(f"Using random seed: {random_seed}")

# Function to read data from xlsx files
def load_data_from_directory(directory_path):
    data_list = []
    for file_name in os.listdir(directory_path):
        if file_name.endswith('.xlsm') and not file_name.startswith('~$'):
            file_path = os.path.join(directory_path, file_name)
            print(f"Reading file: {file_path}")
            try:
                data = load_data_from_file(file_path)
                if all([sheet in data and not data[sheet].empty for sheet in ['Height', 'AntPost', 'LR']]):
                    data_list.append(data)
                else:
                    print(f"Skipping file {file_path} due to missing sheets.")
            except Exception as e:
                print(f"Error reading {file_path}: {e}")
    return data_list

def load_data_from_file(file_path):
    sheets = ['Height', 'AntPost', 'LR']
    data = {}
    try:
        xls = pd.ExcelFile(file_path, engine='openpyxl')
        for sheet in sheets:
            if sheet in xls.sheet_names:
                df = pd.read_excel(file_path, sheet_name=sheet, engine='openpyxl')
                data[sheet] = df
            else:
                data[sheet] = pd.DataFrame()
    except Exception as e:
        print(f"Error loading file {file_path}: {e}")
        data = {sheet: pd.DataFrame() for sheet in sheets}
    return data

# Function to calculate index - does not read A0, A1
def calculate_index(data, indices1, indices2, indices3):
    if data['Height'].empty or data['AntPost'].empty or data['LR'].empty:
        return np.nan

    axial_levels_lr = data['LR'].loc[indices1, 'Level1'].values
    axial_levels_antpost = data['AntPost'].loc[indices3, 'Level1'].values

    if np.any(np.isin(axial_levels_lr, [0, 1])) or np.any(np.isin(axial_levels_antpost, [0, 1, 2])):
        return np.nan

    width = data['LR'].loc[59, 'Length']
    height = data['Height'].loc[41, 'Length']
    length = data['AntPost'].loc[158, 'Length']

    points1 = data['LR'].loc[indices1, 'Length'].values
    points2 = data['Height'].loc[indices2, 'Length'].values
    points3 = data['AntPost'].loc[indices3, 'Length'].values

    points1 = points1[~np.isnan(points1)]
    points2 = points2[~np.isnan(points2)]
    points3 = points3[~np.isnan(points3)]

    avg_point1 = np.mean(points1) if len(points1) > 0 else np.nan
    avg_point2 = np.mean(points2) if len(points2) > 0 else np.nan
    avg_point3 = np.mean(points3) if len(points3) > 0 else np.nan

    if np.isnan(avg_point1) or np.isnan(avg_point2) or np.isnan(avg_point3):
        return np.nan

    index = (np.sqrt((((2 * (avg_point1/width))**2) + ((1 * (avg_point2/height))**2) + ((2 * (avg_point3/length))**2)))/(np.sqrt(3)))
    return index

# Data cleaning - IQR, Z, Winsor
def clean_data(data):
    if not data or len(data) == 0:
        return []

    data = np.array(data)

    Q1 = np.percentile(data, 25)
    Q3 = np.percentile(data, 75)
    IQR = Q3 - Q1
    lower_bound = Q1 - 2.5 * IQR
    upper_bound = Q3 + 2.5 * IQR

    data_iqr_filtered = data[(data >= lower_bound) & (data <= upper_bound)]
    if len(data_iqr_filtered) == 0:
        return []

    z_scores = np.abs(stats.zscore(data_iqr_filtered))
    data_z_filtered = data_iqr_filtered[z_scores < 3.5]
    if len(data_z_filtered) == 0:
        return []

    data_winsorized = np.clip(data_z_filtered, lower_bound, upper_bound)

    return data_winsorized.tolist()

print("Loading UC data...")
UC_data = load_data_from_directory(UC_dir)
if len(UC_data) == 0:
    print("No valid UC data files found. Exiting.")
    exit()
print("UC data loaded.")

print("Loading plagio data...")
plagio_data = load_data_from_directory(plagio_dir)
if len(plagio_data) == 0:
    print("No valid plagio data files found. Exiting.")
    exit()
print("Plagio data loaded.")

max_rows = min(len(UC_data[0]['LR']), len(UC_data[0]['Height']), len(UC_data[0]['AntPost']))

def generate_unique_combinations(max_rows, num_combinations):
    seen_combinations = set()
    combinations_flat = []
    while len(combinations_flat) < num_combinations:
        comb1 = random.sample(range(max_rows), 3)
        comb2 = random.sample(range(max_rows), 3)
        comb3 = random.sample(range(max_rows), 3)
        combination = (*comb1, *comb2, *comb3)
        if combination not in seen_combinations:
            seen_combinations.add(combination)
            combinations_flat.append(combination)
    return combinations_flat

print("Generating unique random combinations...")
combinations_flat = generate_unique_combinations(max_rows, 1000)
print("Unique random combinations generated.")

def balanced_objective_function(combination, UC_data, plagio_data):
    comb1, comb2, comb3 = combination[:3], combination[3:6], combination[6:]
    comb1 = np.clip(comb1, 0, max_rows-1).astype(int)
    comb2 = np.clip(comb2, 0, max_rows-1).astype(int)
    comb3 = np.clip(comb3, 0, max_rows-1).astype(int)

    uc_indices = [calculate_index(data, comb1, comb2, comb3) for data in UC_data]
    plagio_indices = [calculate_index(data, comb1, comb2, comb3) for data in plagio_data]

    uc_indices = clean_data([index for index in uc_indices if not np.isnan(index)])
    plagio_indices = clean_data([index for index in plagio_indices if not np.isnan(index)])

    if len(uc_indices) == 0 or len(plagio_indices) == 0:
        print(f"Combination {combination} resulted in empty cleaned data. Skipping.")
        return float('inf')

    y_true = [1] * len(uc_indices) + [0] * len(plagio_indices)
    y_scores = uc_indices + plagio_indices

    fpr, tpr, thresholds = roc_curve(y_true, y_scores)
    sensitivity = tpr[np.argmax(tpr - fpr)]
    specificity = 1 - fpr[np.argmax(tpr - fpr)]
    score = -(sensitivity + specificity - abs(sensitivity - specificity))
    return score

best_combination = None
best_objective_value = float('inf')
start_time = time.time()
print("Starting random optimization...")
for idx, combination in enumerate(combinations_flat):
    print(f"Evaluating combination {idx+1}/{len(combinations_flat)}: {combination}")
    score = balanced_objective_function(combination, UC_data, plagio_data)
    if score < best_objective_value:
        best_combination = combination
        best_objective_value = score

print(f"Best combination from random optimization: {best_combination} with score: {best_objective_value}")

print("Refining with Nelder-Mead optimization...")
result = minimize(balanced_objective_function, best_combination, args=(UC_data, plagio_data), method='Nelder-Mead')
optimized_combination = np.clip(result.x, 0, max_rows-1).astype(int)
end_time = time.time()

print('Optimization completed.')
print('Optimized combination:', optimized_combination)
print('Best balanced score:', result.fun)
print(f'Elapsed time: {end_time - start_time:.2f} seconds')

best_indices1, best_indices2, best_indices3 = optimized_combination[:3], optimized_combination[3:6], optimized_combination[6:]

def get_cell_references(sheet_name, indices):
    return [f"L{index + 2}" for index in indices]

LR_references = get_cell_references('LR', best_indices1)
Height_references = get_cell_references('Height', best_indices2)
AntPost_references = get_cell_references('AntPost', best_indices3)

print("Cell references to use in the index equation:")
print(f"Points from LR sheet: {LR_references}")
print(f"Points from Height sheet: {Height_references}")
print(f"Points from AntPost sheet: {AntPost_references}")

print("Calculating optimized index values for UC population...")
UC_indices = [calculate_index(data, best_indices1, best_indices2, best_indices3) for data in UC_data]
print("Calculating optimized index values for plagio population...")
plagio_indices = [calculate_index(data, best_indices1, best_indices2, best_indices3) for data in plagio_data]

UC_indices_clean = clean_data([val for val in UC_indices if not np.isnan(val)])
plagio_indices_clean = clean_data([val for val in plagio_indices if not np.isnan(val)])

combined_indices_cleaned = UC_indices_clean + plagio_indices_clean
y_true_cleaned = [1] * len(UC_indices_clean) + [0] * len(plagio_indices_clean)

print("Calculating Sensitivity, Specificity, and optimal cutoff using cleaned data...")
fpr, tpr, thresholds = roc_curve(y_true_cleaned, combined_indices_cleaned)
sensitivity = tpr[np.argmax(tpr - fpr)]
specificity = 1 - fpr[np.argmax(tpr - fpr)]
optimal_idx = np.argmax(tpr - fpr)
optimal_threshold = thresholds[optimal_idx]

print(f"Sensitivity: {sensitivity}")
print(f"Specificity: {specificity}")
print(f"Optimal Cutoff Value: {optimal_threshold}")

print("Plotting histograms with cleaned data...")
plt.figure(figsize=(10, 6))
plt.hist(UC_indices_clean, bins=20, alpha=0.5, label='UC Population', color='blue', density=True)
plt.hist(plagio_indices_clean, bins=20, alpha=0.5, label='Plagio Population', color='orange', density=True)
plt.axvline(x=optimal_threshold, color='red', linestyle='--', label=f'Optimal Cutoff: {optimal_threshold:.2f}')
plt.xlabel('Index Values')
plt.ylabel('Density')
plt.title('Histogram of Index Values for Cleaned Data Population')
plt.legend()
plt.show()
