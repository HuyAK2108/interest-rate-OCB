import msvcrt
import requests
from bs4 import BeautifulSoup
import pandas as pd
import time

today = time.strftime("%d-%m-%Y")
path = r'D:/Lãi suất OCB'
file_name = f"lãi suất ngày {today}"

url = "https://www.ocb.com.vn/vi/cong-cu/lai-suat"
page = requests.get(url)
soup = BeautifulSoup(page.text, 'html.parser')

# Find table
table = soup.find_all('table')[0]

# Find titles Kỳ hạn Tiền gửi có kỳ hạn Tiết kiệm thông thường Tiết kiệm Online
world_titles = table.find_all('tr')
world_table_titles = [title.text.strip() for title in world_titles[0]]

# Display titles
df = pd.DataFrame(columns=world_table_titles[0:4])

# Find data
column_data = table.find_all('tr')
for row in column_data[1:]:
    row_data = row.find_all('td')
    individual_row_data = [data.text.strip() for data in row_data]
    for individual_row_data in individual_row_data[1:]:
        individual_row_data = float(individual_row_data)
    length = len(df)
    df.loc[length] = individual_row_data

# Print Excel daily
df.to_csv(f'{path}/{file_name}.csv', index=False, encoding='utf-8-sig')
df.to_csv(f'log/{file_name}.csv', index=False, encoding='utf-8-sig')

# Print Excel summary
ret = "OK"
try:
    df = pd.read_excel(f'{path}/summary.xlsx')
except:
    df.to_excel(f'{path}/summary.xlsx', index=False, encoding='utf-8-sig', sheet_name=today)
    ret = "NG"

if ret == "OK":
    with pd.ExcelWriter(f'{path}/summary.xlsx', mode = 'a', engine='openpyxl',if_sheet_exists="overlay") as writer:
        df.to_excel(writer, sheet_name=today, index=False)

print(df)
msvcrt.getch()
