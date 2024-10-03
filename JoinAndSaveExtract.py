from config import  fileNameExel
from google_drive_operations import  save_dataframe_to_google_sheet,getFileFromGoogle
import pandas as pd
import logging


        
def EDA(creds,folder_id,spreadsheets_id):
    
    try:
    # Supondo que getFileFromGoogle retorna um dicionário {nome_do_arquivo: DataFrame}
        dataframes_dict = getFileFromGoogle(creds, folder_id, fileNameExel)
        
        if not isinstance(dataframes_dict, dict):
            raise ValueError("A função getFileFromGoogle deve retornar um dicionário de DataFrames.")
        
        # Lista para armazenar os DataFrames processados
        processed_dataframes = []
        
        # Iterar sobre cada DataFrame no dicionário e filtrar as colunas desejadas
        for df_name, df in dataframes_dict.items():
            if not isinstance(df, pd.DataFrame):
                raise ValueError(f"O valor associado a '{df_name}' não é um pd.DataFrame.")
            
            # Verificar se as colunas necessárias estão no DataFrame
            missing_cols = [col for col in ["Period/Segment", "Requirement", "Receipts", "Available quantity"] if col not in df.columns]
            if missing_cols:
                raise ValueError(f"O DataFrame '{df_name}' está faltando as colunas necessárias: {', '.join(missing_cols)}")
        
            # Selecionar as colunas desejadas
            df["Material"] = df_name[5:-5]
            df = df[["Material","Period/Segment", "Requirement", "Receipts", "Available quantity"]]
            # Adicionar o DataFrame processado à lista
            processed_dataframes.append(df)
        
        # Concatenar todos os DataFrames processados
        result = pd.concat(processed_dataframes)
        print(result.head(20))
        # Salvar o DataFrame concatenado em uma planilha do Google Sheets
        save_dataframe_to_google_sheet("Week_Material", f'A1:E', result, creds, "Alex_file", spreadsheets_id)
        
        print('Os DataFrames foram concatenados e salvos com sucesso.')
        
    except Exception as e:
        print(f'Erro ao processar os arquivos: {e}')
        raise