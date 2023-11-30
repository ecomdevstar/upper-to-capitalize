import pandas as pd
import re

df = pd.read_excel('Export_2023-11-28_195754.xlsx')

df['Option1 Value'] = df['Option1 Value'].str.lower()
df['Option1 Value'] = df['Option1 Value'].str.title()

def contains_numbers(value):
    return bool(re.search(r'\d', str(value)))

# Apply conversion to numeric values only
df['Option2 Value'] = df['Option2 Value'].apply(lambda x: pd.to_numeric(x, errors='coerce') if contains_numbers(x) else x)

sort_columns = ['Title', 'Option2 Value']
df = df.sort_values(sort_columns)

df['Variant Position'] = df.groupby('Title').cumcount() + 1

df.to_excel('output.xlsx', index=False)
