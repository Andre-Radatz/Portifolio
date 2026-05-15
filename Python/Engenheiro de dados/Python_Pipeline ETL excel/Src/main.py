import pandas as pd
import os
import glob

folder_path = 'src\\data\\raw'

excel_files = glob.glob(os.path.join(folder_path , '*.xlsx'))

if not excel_files:
    print("nenhum arquivo encontrado")
else:

    dfs = []

for excel_file in excel_files:
    try:
        df_Temp = pd.read_excel(excel_file)

        file_name = os.path.basename(excel_file)

        df_Temp['filename'] = file_name

        if 'brasil' in file_name.lower():
            df_Temp ['location'] = 'br'

        elif 'france' in file_name.lower():
            df_Temp ['location'] = 'fr'

        elif 'italian' in file_name.lower():
            df_Temp ['location'] = 'it'
        
        df_Temp['campaign'] = df_Temp['utm_link'].str.extract(r'utm_campaign=(.*)')

        dfs.append(df_Temp)

        print(df_Temp)
    except Exception as e:
        print (f"erro no arquivo {excel_file} : {e} ")

if dfs:
    result = pd.concat(dfs, ignore_index=True)

    outpout_file = os.path.join('Src', 'Data', 'Ready', 'clean.xlsx')

    writer = pd.ExcelWriter(outpout_file, engine='xlsxwriter')

    result.to_excel(writer, index=False)

    writer._save()
else:
    print('nenhum dado encontrado')