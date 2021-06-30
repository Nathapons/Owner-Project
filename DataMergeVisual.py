import matplotlib.pyplot as plt
import pandas as pd
import numpy as np
import statistics

csv_file = "D:\\Nathapon_KeepFolder\\0.My work\\01.IoT\\08.P'Muse\\SERVER\\20210627\\RESULT\\ACT_RESULT.CSV"
df = pd.read_csv(csv_file)

datetime_lists = []
avg1_lists = []
tooltips = []
max_row = list(df.shape)[0]

count = 0
for row in range(max_row):
    avg1 = float(df['CPU System'][row])
    date = str(df['Date'][row])
    time = str(df['Time'][row])
    datetime = date + " " + time
    tooltip = "Date: " + datetime + " Values: " + str(avg1)

    avg1_lists.append(avg1)
    datetime_lists.append(count)
    tooltips.append(tooltip)

    count += 1
    # if count >= 10:
    #     break

# Plot Graph
plt.plot(datetime_lists, avg1_lists, linestyle='-', marker='.', color='black')
mean = statistics.mean(avg1_lists)
plt.axhline(mean, color='green')

# Graph view Setting
max_val = int(max(avg1_lists))
min_val = int(min(avg1_lists))
ymarg = (max_val - min_val) * plt.margins()[1]
plt.yticks(np.arange(min_val, max_val, 2))
title_setting = {'size':20,'color':'blue'}
plt.ylabel('Load Average 1m')
plt.xlabel('Datetime', rotation=0)
plt.title('Load Averge 1m', title_setting)

plt.show()