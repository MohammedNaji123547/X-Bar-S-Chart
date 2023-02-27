import matplotlib.pyplot as plt
from openpyxl.drawing.image import Image
import io
import sys
from datetime import datetime
import pandas as pd
from random_sampling import get_random_sample
from openpyxl.chart import (line_chart,reference,series)

from biasing_constrains_appendix import *

print(sys.argv)

excel_file = 'test.xlsx'

random_sample = get_random_sample(excel_file)

result_df = pd.DataFrame()

result_df['x_bar'] = random_sample.mean(axis=0)
result_df['x_bar_bar'] = result_df['x_bar'].mean(axis=0)
result_df['s_bar'] = random_sample.std(axis=0)
result_df['s_bar_bar'] = result_df['s_bar'].mean(axis=0)


currentDate = datetime.now()
timestamp = datetime.timestamp(currentDate)

with pd.ExcelWriter('D:/' + str(timestamp) + '.xlsx') as writer:
    random_sample.to_excel(writer, sheet_name='randomized-sample')
    result_df.to_excel(writer, sheet_name='statistics')

random_sample_columns = random_sample.columns
random_sample_columns_count = len(random_sample_columns) - 1

if 10<random_sample_columns_count<15:
    biasing_constrains = {"A3": 0.975, "B3": 0.284, "B4": 1.716}

elif 15<=random_sample_columns_count<25:
     biasing_constrains= {"A3":0.789,"B3":0.428, "B4":1.572}

elif random_sample_columns_count >= 25:
     biasing_constrains= {"A3":0.606,"B3":0.565, "B4":1.435}

s_bar_bar = 's_bar_bar'
x_bar_bar = 'x_bar_bar'
UCLs = result_df[s_bar_bar].mean() + biasing_constrains["B3"] * result_df['s_bar_bar'].mean()
LCLs = result_df[s_bar_bar].mean() - biasing_constrains["B4"] * result_df['s_bar_bar'].mean()
UCLx = result_df[x_bar_bar].mean() + biasing_constrains["A3"] * result_df['s_bar_bar'].mean()
LCLx = result_df[x_bar_bar].mean() - biasing_constrains["A3"] * result_df['s_bar_bar'].mean()

result_df = result_df.assign(UCLs=UCLs, LCLs=LCLs, UCLx=UCLx, LCLx=LCLx)
print(result_df)
# Create X-bar S-chart using Matplotlib
plt.figure(figsize=(10,6))
plt.plot(result_df.index, result_df['x_bar'], '-o', color='blue', label='X-bar')
plt.plot(result_df.index, result_df['s_bar'], '-o', color='red', label='S')
plt.axhline(y=result_df['x_bar_bar'][0], color='gray', linestyle='--', label='X-bar-bar')
plt.axhline(y=UCLx, color='gray', linestyle='--', label='UCLx')
plt.axhline(y=LCLx, color='gray', linestyle='--', label='LCLx')
plt.axhline(y=result_df['s_bar_bar'][0], color='gray', linestyle='-.', label='S-bar')
plt.axhline(y=UCLs, color='gray', linestyle='-', label='UCLs')
plt.axhline(y=LCLs, color='gray', linestyle='-', label='LCL')
print(plt.axhline)
# Save the chart as a PNG image in memory
buf = io.BytesIO()
plt.savefig(buf, format='png')
buf.seek(0)

# Insert the image into an Excel file
from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.drawing.image import Image

wb = Workbook()
ws = wb.active
img = Image(buf)
ws.add_image(img, 'A1')

# Save the Excel file
wb.save('D:/' + str(timestamp) + '.xlsx')

currentDate = datetime.now()
timestamp = datetime.timestamp(currentDate)

with pd.ExcelWriter('D:/' + str(timestamp) + '.xlsx') as writer:
    random_sample.to_excel(writer, sheet_name='randomized-sample')
    result_df.to_excel(writer, sheet_name='statistics')

# TODO: test string input of numerical value
#---------------------------

