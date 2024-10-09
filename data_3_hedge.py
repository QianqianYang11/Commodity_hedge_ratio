import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

data = pd.read_excel('D:\\2023semester\\Commodity_hedge_ratio\\Data.xlsx')
data['Date'] = pd.to_datetime(data['Date'])
data.set_index('Date', inplace=True)
data = data.iloc[1:]
dates = data.index

qt_data = pd.read_csv('D:\\2023semester\\Commodity_hedge_ratio\\vech_t_cov.csv')

cov_matrices = []
for index, row in qt_data.iterrows():
    matrix = np.zeros((3, 3))
    var_cols = ['Var(METAL)', 'Var(ENERGY)', 'Var(AGRICULTURAL)']
    cov_cols = [('METAL', 'ENERGY'), ('METAL', 'AGRICULTURAL'), ('ENERGY', 'AGRICULTURAL')]
    for i in range(3):
        matrix[i, i] = row[var_cols[i]]
    for i, (col1, col2) in enumerate(cov_cols):
        cov_col = f'Cov({col1},{col2})'
        if cov_col in row:
            value = row[cov_col]
            row_idx, col_idx = (0, 1) if i == 0 else (0, 2) if i == 1 else (1, 2)
            matrix[row_idx, col_idx] = value
            matrix[col_idx, row_idx] = value
    cov_matrices.append(matrix)

Qt = np.array(cov_matrices)
Qt = np.transpose(Qt, axes=(1, 2, 0))
T = Qt.shape[2]
hedge_ratios = {
    'Hedge Metal with Energy': Qt[0, 1, :] / Qt[1, 1, :],
    'Hedge Metal with Agricultural': Qt[0, 2, :] / Qt[2, 2, :],
    'Hedge Energy with Metal': Qt[1, 0, :] / Qt[0, 0, :],
    'Hedge Energy with Agricultural': Qt[1, 2, :] / Qt[2, 2, :],
    'Hedge Agricultural with Metal': Qt[2, 0, :] / Qt[0, 0, :],
    'Hedge Agricultural with Energy': Qt[2, 1, :] / Qt[1, 1, :]
}

colors = ['yellow', 'green', 'orange', 'royalblue', 'darkturquoise', 'olive']
markers = ['o', 's', 'x', '^', 'd', 'v']
line_width = 2

y_axis_range = [-0.5, 1.5]

fig, axes = plt.subplots(3, 1, figsize=(15, 8), sharex=True)

# Plot Hedge METAL
axes[0].plot(dates, hedge_ratios['Hedge Metal with Energy'], label='Hedge Metal with Energy', linewidth=line_width, color=colors[0], marker=markers[0], markevery=int(len(dates) / 20))
axes[0].plot(dates, hedge_ratios['Hedge Metal with Agricultural'], label='Hedge Metal with Agricultural', linewidth=line_width, color=colors[1], marker=markers[1], markevery=int(len(dates) / 20))
axes[0].set_ylim(y_axis_range)
# axes[0].set_ylabel('Hedge Ratio', fontsize=14)
axes[0].legend(bbox_to_anchor=(1.05, 1), loc='upper left', fontsize=12)
axes[0].set_title('Hedge Metal', fontsize=18)
axes[0].tick_params(axis='y', labelsize=10)
axes[0].tick_params(axis='x', rotation=45, labelsize=10)

# Plot Hedge ENERGY
axes[1].plot(dates, hedge_ratios['Hedge Energy with Metal'], label='Hedge Energy with Metal', linewidth=line_width, color=colors[2], marker=markers[2], markevery=int(len(dates) / 20))
axes[1].plot(dates, hedge_ratios['Hedge Energy with Agricultural'], label='Hedge Energy with Agricultural', linewidth=line_width, color=colors[3], marker=markers[3], markevery=int(len(dates) / 20))
axes[1].set_ylim(y_axis_range)
# axes[1].set_ylabel('Hedge Ratio', fontsize=14)
axes[1].legend(bbox_to_anchor=(1.05, 1), loc='upper left', fontsize=12)
axes[1].set_title('Hedge Energy', fontsize=18)
axes[1].tick_params(axis='y', labelsize=10)
axes[1].tick_params(axis='x', rotation=45, labelsize=10)

# Plot Hedge AGRICULTURAL
axes[2].plot(dates, hedge_ratios['Hedge Agricultural with Metal'], label='Hedge Agricultural with Metal', linewidth=line_width, color=colors[4], marker=markers[4], markevery=int(len(dates) / 20))
axes[2].plot(dates, hedge_ratios['Hedge Agricultural with Energy'], label='Hedge Agricultural with Energy', linewidth=line_width, color=colors[5], marker=markers[5], markevery=int(len(dates) / 20))
axes[2].set_ylim(y_axis_range)
# axes[2].set_ylabel('Hedge Ratio', fontsize=14)
# axes[2].set_xlabel('Date', fontsize=14)
axes[2].legend(bbox_to_anchor=(1.05, 1), loc='upper left', fontsize=12)
axes[2].set_title('Hedge Agricultural', fontsize=18)
axes[2].tick_params(axis='y', labelsize=10)
axes[2].tick_params(axis='x', rotation=45, labelsize=10)

fig.subplots_adjust(right=0.8, hspace=0.2)
plt.show()

statistics = {
    'Mean': [],
    'Std': [],
    'Min': [],
    'Max': []
}

for key, value in hedge_ratios.items():
    statistics['Mean'].append(np.mean(value))
    statistics['Std'].append(np.std(value))
    statistics['Min'].append(np.min(value))
    statistics['Max'].append(np.max(value))

statistics_df = pd.DataFrame(statistics, index=hedge_ratios.keys())

excel_file_path = 'D:\\2023semester\\Commodity_hedge_ratio\\hedge_ratio.xlsx'

statistics_df.to_excel(excel_file_path)

print("Statistics of hedge ratios saved to:", excel_file_path)


