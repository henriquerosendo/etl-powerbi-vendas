import pyodbc
import pandas as pd
import os
import shutil
import glob


#limpar a tela 
os.system("cls")


#origem
pasta_origem  = "planilhas"
pasta_destino = "planilhas/processado"
pasta_error   = "planilhas/error"


#caso nao exista a pasta cria as pastas de error ou de processados 
os.makedirs(pasta_destino, exist_ok=True) 
os.makedirs(pasta_error, exist_ok=True)


arquivos_csv = glob.glob(os.path.join(pasta_origem, "*.csv"))

if not arquivos_csv:
    print("Nenhum arquivo econtrado na pasta planilhas")
else:
    for arquivo_origem  in arquivos_csv:
        print("Processando os arquivos CSV, arquivo:",arquivo_origem )

        try:
            # Ler CSV com separador ;
            df = pd.read_csv(arquivo_origem, sep=';')

            # Conexão com SQL Server
            connection = (
                "Driver={SQL Server};"
                "Server=RosendoPc\SQLEXPRESS;"
                "Database=AnalyticsDB;"
            )
            conexao = pyodbc.connect(connection)
            cursor = conexao.cursor()
            print("Conexão bem sucedida")

            # Inserir dados do DataFrame no banco
            for index, row in df.iterrows():
                comando = f"""
                INSERT INTO vendas (DataVenda, Produto, Quantidade, PrecoUnitario)
                VALUES ('{row['DataVenda']}', '{row['Produto']}', {row['Quantidade']}, {row['PrecoUnitario']})
                """
                cursor.execute(comando)

            # Salvar alterações
            conexao.commit()
            print("Dados inseridos com sucesso!")

            # Fechar conexão
            cursor.close()
            conexao.close()


            arquivo_destino = os.path.join(pasta_destino, os.path.basename(arquivo_origem))
            shutil.move(arquivo_origem,arquivo_destino)

            print(f"Arquivo movido para '{arquivo_destino}' com sucesso!")


        except Exception as e:
            print("Erro ao processar o arquivo:", e)

            # Mover para pasta de erro
            arquivo_destino = os.path.join(pasta_error, os.path.basename(arquivo_origem))
            shutil.move(arquivo_origem, arquivo_destino)
            print(f"Arquivo movido para '{arquivo_destino}' devido a erro.")