import camelot
import pandas as pd
import os

def processar_pdf_para_excel(pdf_path, excel_path):
    print(f"\n📄 Processando: {pdf_path}")

    try:
        tabelas = camelot.read_pdf(pdf_path, pages='all', flavor='lattice', strip_text='\n')

        if not tabelas or tabelas.n == 0:
            print("⚠️ Nenhuma tabela encontrada. Tentando com flavor='stream'...")
            tabelas = camelot.read_pdf(pdf_path, pages='all', flavor='stream', strip_text='\n')
            if not tabelas or tabelas.n == 0:
                print("⚠️ Nenhuma tabela encontrada com nenhum dos flavors.")
                return

        print(f"🔍 Tabelas detectadas: {len(tabelas)}")

        lista_dfs = []
        max_cols_global = 0

        for tabela in tabelas:
            df = tabela.df.dropna(axis=1, how='all')
            max_cols_global = max(max_cols_global, df.shape[1])
            lista_dfs.append(df)

        df_final_lista = []
        for i, df in enumerate(lista_dfs):
            df_padronizado = df.reindex(columns=range(max_cols_global))
            df_final_lista.append(df_padronizado)

        df_total = pd.concat(df_final_lista, ignore_index=True)
        df_total.to_excel(excel_path, index=False)
        print(f"✅ Arquivo salvo em: {excel_path}")

    except Exception as e:
        print(f"❌ Erro ao processar {pdf_path}: {e}")

# 👇 Pasta com os arquivos PDF
pasta_pdf = "./PDF"  # <--- Altere aqui
pasta_saida = os.path.join(pasta_pdf, "tabelas_extraidas")
os.makedirs(pasta_saida, exist_ok=True)

# 👇 Loop para processar todos os PDFs da pasta
for nome_arquivo in os.listdir(pasta_pdf):
    if nome_arquivo.lower().endswith(".pdf"):
        caminho_pdf = os.path.join(pasta_pdf, nome_arquivo)
        nome_base = os.path.splitext(nome_arquivo)[0]
        caminho_excel = os.path.join(pasta_saida, f"{nome_base}.xlsx")
        processar_pdf_para_excel(caminho_pdf, caminho_excel)
