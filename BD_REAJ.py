import pandas as pd
from datetime import datetime

# Load the uploaded Excel file
input_file_path = 'BD_REAJ.xlsx'  # Use your file path here
data = pd.read_excel(input_file_path)

# Creating a new DataFrame to store the result
expanded_data = []

# Loop through each unique project (Contratante, Grupo, Data_Base)
for (contratante, grupo, data_base), project_data in data.groupby(['Contratante', 'Grupo', 'Data_Base']):
    project_data = project_data.sort_values(by='Mês_Ano')
    
    # Loop through each row in the project-specific data
    for index, row in project_data.iterrows():
        mes_ano = pd.to_datetime(row['Mês_Ano'])
        reajuste = row['R%']
        
        # Generate monthly rows starting from the given Mês_Ano until the next defined period or end
        if index < len(project_data) - 1:
            next_mes_ano = pd.to_datetime(project_data.iloc[index + 1]['Mês_Ano'])
        else:
            next_mes_ano = mes_ano + pd.DateOffset(years=3)
        
        start_date = mes_ano
        while start_date < next_mes_ano:
            expanded_data.append([contratante, grupo, data_base, start_date.strftime('%m/%Y'), reajuste])
            start_date = start_date + pd.DateOffset(months=1)

# Create a new DataFrame from the expanded data
result_df = pd.DataFrame(expanded_data, columns=['Contratante', 'Grupo', 'Data_Base', 'Mês_Ano', 'R%'])

# Remove duplicates, keeping only the rows with the most recent 'R%' value for each 'Contratante', 'Grupo', 'Data_Base', and 'Mês_Ano'
result_df['Mês_Ano'] = pd.to_datetime(result_df['Mês_Ano'], format='%m/%Y')
result_df.sort_values(by=['Contratante', 'Grupo', 'Data_Base', 'Mês_Ano', 'R%'], ascending=[True, True, True, True, False], inplace=True)
result_df = result_df.drop_duplicates(subset=['Contratante', 'Grupo', 'Data_Base', 'Mês_Ano'], keep='first')

# Save the new DataFrame to an Excel file
output_file_path = 'BD_REAJ_expanded_3years.xlsx'  # Use your desired output path
result_df.to_excel(output_file_path, index=False)

print(f"Arquivo salvo em: {output_file_path}")
