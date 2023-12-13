import pandas as pd

# Define the sort order
sort_order = {
    '提高班': 1,
    '中级班': 2,
    '初中级班': 3,
    '初级班': 4
}

# Load the data
file_path = 'input.xlsx'  # Replace with your file path
df = pd.read_excel(file_path)

# Map the class levels to their sort order and sort
df['sort_key'] = df['请选择自己的班級'].map(sort_order)
df_sorted = df.sort_values(by='sort_key')

# Function to label groups (Group A, Group B, etc.)
def label_group(number):
    return 'Group' + chr(65 + number)

# Divide the DataFrame into groups of 8 and write to separate text files with GB18030 encoding
for i in range(0, len(df_sorted), 8):
    group_label = label_group(i // 8)
    file_name = f'{group_label}.txt'
    with open(file_name, 'w', encoding='GB18030') as file:
        names = df_sorted.iloc[i:i + 8]['华语姓名'].tolist()
        for name in names:
            file.write(f'{name}\n')