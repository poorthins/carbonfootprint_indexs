import pandas as pd
from openpyxl import load_workbook

# 指定檔案路徑
file_path = '碳足跡\\工程\\5.xlsx'  # 請替換成您的檔案路徑
carbon_file_path = "碳足跡\\係數\\新的\\5_index.xlsx"
carbon_data = pd.read_excel(carbon_file_path)

# Part 1
# 讀取資源統計表
resource_data = pd.read_excel(file_path, sheet_name='資源統計表')

# 建立編碼到碳足跡的映射
code_to_carbon_map = dict(zip(carbon_data['編碼'], carbon_data['碳係數']))

# 更新資源統計表中的碳足跡數據
resource_data['碳係數'] = resource_data['編碼'].map(code_to_carbon_map)

# 將更新後的資源統計表保存回原檔案
with pd.ExcelWriter(file_path, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
    resource_data.to_excel(writer, sheet_name='資源統計表', index=False)

# 使用Pandas讀取Excel檔案
excel_data = pd.ExcelFile(file_path)
# 讀取'資源統計表'和'單價分析表'
resource_statistics_sheet = excel_data.parse('資源統計表')
unit_price_analysis_sheet = excel_data.parse('單價分析表')

carbon_data = resource_data.drop(resource_data.columns[4], axis=1)
unit_price_analysis_sheet = unit_price_analysis_sheet.drop(columns=[
    unit_price_analysis_sheet.columns[7],
    unit_price_analysis_sheet.columns[8],
    unit_price_analysis_sheet.columns[9],
    #unit_price_analysis_sheet.columns[10]
])

# 合併'單價分析表'與'資源統計表'的數據
merged_unit_price_analysis_sheet = unit_price_analysis_sheet.merge(
    resource_statistics_sheet[['編碼', '碳係數']],
    on='編碼',
    how='left'
)

# 將修改後的'單價分析表'保存回原Excel檔案
with pd.ExcelWriter(file_path, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
    merged_unit_price_analysis_sheet.to_excel(writer, sheet_name='單價分析表', index=False)

# Part 2 sheet
# 讀取Excel檔案中的「單價分析表」工作表  
data_for_process = pd.read_excel(file_path, sheet_name='單價分析表')

total_carbon_per_item = 0
total_volume_per_item = 0
item_start_index = None
appendix_data = []

# 增加2個欄位
data_for_process['目前有連結'] = False
data_for_process['工項單位碳排'] = ''

df_row_of_workitems = data_for_process[data_for_process['項次'].notna()]
list_of_workitems_index = data_for_process[data_for_process['項次'].notna()].index.to_list()
list_of_workitems = data_for_process[data_for_process['項次'].notna()]['項次'].to_list()

interval = [list_of_workitems_index[i] - list_of_workitems_index[i-1] for i in range(1, len(list_of_workitems_index))]
interval.append(len(data_for_process)-list_of_workitems_index[-1])

df_row_of_appendix = data_for_process[data_for_process['項次']=='appendix']
list_of_appendix_index = data_for_process[data_for_process['項次']=='appendix'].index.to_list()
list_of_appendix_no = df_row_of_appendix['編碼'].to_list()
package = data_for_process[(data_for_process['編碼'].isin(list_of_appendix_no)) & (data_for_process['項次']!='appendix') ]

data_for_process.loc[package.index, '目前有連結'] = True

def run_check_children():
    for num, index in enumerate(list_of_workitems_index):
        aaa = True if True in data_for_process['目前有連結'].iloc[index+1:index+interval[num]].values else False
        data_for_process.loc[index, '目前有連結'] = aaa

run_check_children()

# Calculate the max consective appendix
def max_consecutive_appendix(names):
    max_count = 0
    current_count = 0

    for name in names:
        if name=='appendix':  
            current_count += 1
            max_count = max(max_count, current_count)
        else:
            current_count = 0

    return max_count

max_con = max_consecutive_appendix(list_of_workitems)

def cal(i, workitem_index):
    
    df_workitem = data_for_process.iloc[workitem_index:workitem_index+interval[i]]
    
    carbon_footprint = 0
    total_carbon_per_item = 0
    total_volume_per_item = 0
    item_start_index = None
    
    for index, row in df_workitem.iterrows():
        
        if pd.notna(row['項次']):
            item_start_index = index

        if row['名稱'] != "合計":
            volume = 0.0 if row['用量'] == 'NP' else float(row['用量']) if pd.notna(row['用量']) else 0.0
            carbon_coef = 0.0 if row['碳係數'] == 'NP' else float(row['碳係數']) if pd.notna(row['碳係數']) else 0.0
            total_carbon_per_item += volume * carbon_coef
        else:
            total_volume_per_item = row['用量']

    if item_start_index is not None and total_volume_per_item != 0:
        carbon_footprint = total_carbon_per_item / total_volume_per_item
        data_for_process.at[item_start_index, '工項單位碳排'] = carbon_footprint


def copy_back(no, L):
    data_for_process.loc[data_for_process['編碼'] == no, '碳係數'] = L
    data_for_process.loc[data_for_process['編碼'] == no, '目前有連結'] = False


for t in range(max_con+1):
    for i, workitem_index in enumerate(list_of_workitems_index): 
        if data_for_process.loc[workitem_index, '目前有連結']==False and data_for_process.loc[workitem_index, '工項單位碳排']=='':
            cal(i, workitem_index)
            if data_for_process.loc[workitem_index, '項次'] == 'appendix':
                no = data_for_process.loc[workitem_index, '編碼']
                L = data_for_process.loc[workitem_index, '工項單位碳排']
                copy_back(no, L)
    run_check_children()

# 將計算完成的DataFrame保存回原Excel檔案的「單價分析表」工作表
with pd.ExcelWriter(file_path, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
    data_for_process.to_excel(writer, sheet_name='單價分析表', index=False)

# Part 3 sheet

###合回去標單詳細表
unit_price_analysis_sheet = pd.read_excel(file_path, sheet_name='單價分析表')
bid_detail_sheet = pd.read_excel(file_path, sheet_name='標單詳細表')

bid_detail_sheet = bid_detail_sheet.drop(columns=[
    bid_detail_sheet.columns[6],
    bid_detail_sheet.columns[7],
])

# 過濾出「單價分析表」中「項次」非空格的行
filtered_unit_price_analysis_sheet = unit_price_analysis_sheet[unit_price_analysis_sheet['項次'].notna()]

# 建立編碼到L值的映射
code_to_L_map = dict(zip(filtered_unit_price_analysis_sheet['編碼'], filtered_unit_price_analysis_sheet['工項單位碳排']))

# 在「標單詳細表」中添加一個新列，用於存儲根據編碼匹配到的L值
bid_detail_sheet['工項單位碳排'] = bid_detail_sheet['項次'].map(code_to_L_map)

# 將修改後的「標單詳細表」保存回原Excel檔案
with pd.ExcelWriter(file_path, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
    bid_detail_sheet.to_excel(writer, sheet_name='標單詳細表', index=False)


bid_detail_sheet = pd.read_excel(file_path, sheet_name='標單詳細表')
resource_statistics_sheet = pd.read_excel(file_path, sheet_name='資源統計表')

# 為了方便匹配，建立一個從編碼到碳係數的映射字典
carbon_dict = dict(zip(resource_statistics_sheet['編碼'], resource_statistics_sheet['碳係數']))

for index, row in bid_detail_sheet.iterrows():
    # 檢查標單詳細表中的項次是否存在於carbon_dict中
    if row['項次'] in carbon_dict:
        # 如果存在，將對應的碳係數填入L列
        bid_detail_sheet.at[index, '工項單位碳排'] = carbon_dict[row['項次']]

# 讀取Excel檔案中的「單價分析表」和「標單詳細表」工作表
bid_detail_sheet['工項總碳排'] = bid_detail_sheet['用量'] * bid_detail_sheet['工項單位碳排']

# 將填充後的標單詳細表保存回原Excel檔案
with pd.ExcelWriter(file_path, mode='a', engine='openpyxl', if_sheet_exists='replace') as writer:
    bid_detail_sheet.to_excel(writer, sheet_name='標單詳細表', index=False)

total_carbon = {}
total= 0
# 使用 iterrows 遍历 bid_detail_sheet
for index, row in bid_detail_sheet.iterrows():

    item = row['工項']
    carbon_value = pd.to_numeric(row['工項總碳排'], errors='coerce')
    if pd.isna(carbon_value):
        continue
    if item not in total_carbon:
        total_carbon[item] = carbon_value
        total += carbon_value
    else:
        total_carbon[item] += carbon_value
        total += carbon_value
for item, carbon_value in total_carbon.items():
    print(f"{item}: {carbon_value}")
print(total)