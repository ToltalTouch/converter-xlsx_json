import pandas as pd
import logging
import glob
import os

pasta = os.path.dirname(os.path.abspath(__file__))
arquivos = glob.glob(os.path.join(pasta, "*.xlsx"))

for arquivo in arquivos:
    logging.info(f"Lendo arquivo: {arquivo}")

    df = pd.read_excel(arquivo, sheet_name=0)
    
    for col in df.select_dtypes(include=["datetime", "datetimetz"]).columns:
        df[col] = df[col].dt.strftime("%d/%m/%Y")

    for col in df.select_dtypes(include=["number"]).columns:
        df[col] = df[col].apply(
            lambda x: str(int(x)) if pd.notnull(x) and float(x).is_integer() else str(x)
        )

    json_data = df.to_json(orient="records",
                        force_ascii=False,
                        indent=4
                        )

    json_data = json_data.replace("\\/", "/")

    nome_json = os.path.splitext(os.path.basename(arquivo))[0] + ".json"
    caminho_json = os.path.join(pasta, nome_json)

    with open(caminho_json, "w", encoding="utf-8") as file:
        file.write(json_data)
        
    print("JSON gerado com sucesso")