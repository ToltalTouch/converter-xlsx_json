import pandas as pd
import logging
import glob
import sys
import os
from datetime import datetime


class ProcessadorExcel:
    def __init__(self):
        logging.basicConfig(level=logging.INFO, format='%(asctime)s - %(levelname)s - %(message)s')
        if getattr(sys, 'frozen', False):
            self.pasta = os.path.dirname(sys.executable)
        else:
            self.pasta = os.path.dirname(os.path.abspath(__file__))
            
        self.arquivos = glob.glob(os.path.join(self.pasta, "*.xlsx"))

    
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
        if not self.arquivos:
            logging.warning("Nenhum arquivo Excel (.xlsx) encontrado na self.pasta atual.")
        else:
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
                caminho_json = os.path.join(self.pasta, nome_json)

                with open(caminho_json, "w", encoding="utf-8") as file:
                    file.write(json_data)
                logging.info("JSON criado com sucesso")
            
if __name__ == "__main__":
    processador = ProcessadorExcel()
    processador.main()
    input("Pressione Enter para sair...")