import pandas as pd
import logging
import glob
import os
from datetime import datetime

logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')

pasta = os.path.dirname(os.path.abspath(__file__))
arquivos = glob.glob(os.path.join(pasta, 'excel', "*.xlsx"))

class ProcessadorExcel:
    def __init__(self, arquivos):
        self.arquivos = arquivos
    
    def processar_linha(self, row):
        for col in row.index:
            valor = row[col]
            if pd.notnull(valor):
                if isinstance(valor, (pd.Timestamp, datetime)):
                    row[col] = valor.strftime("%d/%m/%Y")
                elif isinstance(valor, (int, float)) and not isinstance(valor, bool):
                    row[col] = str(int(valor)) if float(valor).is_integer() else str(valor)
        return row

    def main(self):
        for arquivo in self.arquivos:
            logging.info(f"Lendo arquivo: {arquivo}")

            df = pd.read_excel(arquivo)
            df = df.apply(self.processar_linha, axis=1)
            
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
            caminho_json = os.path.join(pasta, 'json', nome_json)

            with open(caminho_json, "w", encoding="utf-8") as file:
                file.write(json_data)
                
            print("JSON gerado com sucesso")
            
if __name__ == "__main__":
    if not arquivos:
        logging.warning("Nenhum arquivo Excel (.xlsx) encontrado na pasta atual.")
    else:
        processador = ProcessadorExcel(arquivos)
        processador.main()