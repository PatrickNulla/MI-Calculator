import pandas as pd
import numpy as np
import json

# Read the Excel file into a pandas dataframe
df = pd.read_excel('Book1.xlsx')

# Read the scope JSON file into a Python dictionary
with open('scope.json') as f:
    scope = json.load(f)

# Get the acceptable scopes from the dictionary
acceptable_scopes = scope['columnScope']

# Filter the dataframe to only include rows where the 'Scope' column equals an acceptable scope
df = df[df['Scope'].isin(acceptable_scopes)]

# Filter the dataframe to exclude rows where the 'Project' column contains '.tests'
df = df[~df['Project'].str.contains('.Tests')]

# Filter the dataframe to only include the 'Project' and 'Maintainability Index' columns
df = df[['Scope', 'Project', 'Maintainability Index']]

# Read the filterOut JSON file into a Python list
with open('filter.json') as f:
    filter_out = json.load(f)['filteredProjects']
    filter_out = [p.replace('\\', '/') for p in filter_out]
    print(filter_out)

# Filter the dataframe to exclude the projects specified in the filterOut list
df['Project'] = df['Project'].str.replace('\\', '/')
df.loc[df['Project'].isin(filter_out), 'Maintainability Index'] = np.nan

# Sort the result dataframe by the index (the project names) in alphabetical order
df = df.sort_index()
df.rename(columns={'_2': 'MaintainabilityIndex'}, inplace=True)

# Set the maximum number of rows and columns to display
pd.options.display.max_rows = None
pd.options.display.max_columns = None

# Get the maximum length of the project names
max_length_scope = max([len(p) for p in df['Scope']])
max_length_project = max([len(p) for p in df['Project']]) + 1

# Get the maximum index number
max_index = len(df)
max_length_index = len(str(max_index)) + 1

# Loop through the rows in the dataframe
for i, row in enumerate(df.itertuples(), 1):
    # Print the counter, the project name, and the maintainability index, with equal spacing
    print(f'[{i:>{max_length_index}}]  {row.Scope:<{max_length_scope}} - {row.Project:<{max_length_project}}: {row[3]:.2f}')


# Compute and print the overall average of the 'Maintainability Index' column
overall_average = df['Maintainability Index'].mean(skipna=True)
print(f'\nOverall average: {overall_average:.2f}')

# Compute and print the total count of projects
total_count = df.shape[0]
print(f'Total count of projects: {total_count}')

# Compute and print the total count of projects with NaN value
nan_count = df['Maintainability Index'].isna().sum()
print(f'Total count of projects with NaN value: {nan_count}')

# Compute and print the total count of projects excluding projects with NaN value
excluding_nan_count = total_count - nan_count
print(f'Total count of projects excluding projects with NaN value: {excluding_nan_count}')

# Write the dataframe to a CSV file
df.to_csv('result.csv', index=False)

# Write the 'Maintainability Index' column to a CSV file
df['Maintainability Index'].to_csv('maintainability_index_values_only.csv', header=False, index=False)
