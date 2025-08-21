import pandas as pd
import logging
import glob
import os

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

pasta = os.path.dirname(os.path.abspath(__file__))
arquivos = glob.glob(os.path.join(pasta, "*.xlsx"))

def processar_linha(row):
    for col in row.index:
        valor =row[col]
        if pd.notnull(valor):
            if pd.notnull(valor):
                if pd.api.types.is_datetime64_any_dtype(df[col]):
                    row[col] = valor.strftime("%d/%m/%Y")
                elif pd.api.types.is_numeric_dtype(df[col]):
                    row[col] = str(int(valor)) if float(valor).is_integer() else str(valor)
    return row

def main():
    for arquivo in arquivos:
        logging.info(f"Lendo arquivo: {arquivo}")

        df = pd.read_excel(arquivo, sheet_name=0)
        df = df.apply(processar_linha, axis=1)
        
        for i in range(len(df)):
            try:
                logging.info(f"Linha {i} concluida")
            except Exception as e:
                logging.error(f"Erro ao processar linha {i}: {e}")

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
        
if __name__ == "__main__":
    main()