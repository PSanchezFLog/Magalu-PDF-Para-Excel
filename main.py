import camelot
import pandas as pd
import os

def processar_pdf_para_excel(pdf_path, excel_path):
    """
    Processa um arquivo PDF, extrai tabelas e as salva em um arquivo Excel.

    Args:
        pdf_path (str): O caminho para o arquivo PDF de entrada.
        excel_path (str): O caminho para o arquivo Excel de saída.
    """
    print(f"\n📄 Processando: {pdf_path}")

    try:
        # Tenta usar 'lattice' primeiro, que é melhor para tabelas com linhas visíveis.
        # Se não funcionar bem, pode-se tentar 'stream' com ajustes.
        # Ajuste 'table_areas' se souber a área exata da tabela em cada página.
        tabelas = camelot.read_pdf(pdf_path, pages='all', flavor='lattice', strip_text='\n')

        if not tabelas:
            print("⚠️ Nenhuma tabela encontrada. Tentando com flavor='stream'...")
            # Se 'lattice' não encontrar nada, tenta 'stream'
            tabelas = camelot.read_pdf(pdf_path, pages='all', flavor='stream', strip_text='\n')
            if not tabelas:
                print("⚠️ Nenhuma tabela encontrada com nenhum dos flavors.")
                return

        print(f"🔍 Tabelas detectadas: {len(tabelas)}")

        lista_dfs = []
        max_cols_global = 0

        # Primeira passagem para determinar o número máximo de colunas
        # e coletar os DataFrames.
        for i, tabela in enumerate(tabelas):
            df = tabela.df
            # O Camelot pode adicionar colunas extras vazias. Remove-as.
            df = df.dropna(axis=1, how='all')
            if df.shape[1] > max_cols_global:
                max_cols_global = df.shape[1]
            lista_dfs.append(df)

        print(f"Número máximo de colunas detectadas em todas as tabelas: {max_cols_global}")

        # Segunda passagem para padronizar o número de colunas de cada DataFrame
        # e lidar com os cabeçalhos.
        df_final_lista = []
        primeiro_df_processado = False

        for i, df in enumerate(lista_dfs):
            # Padroniza o número de colunas para o máximo encontrado
            # Isso garante que todos os DFs tenham o mesmo número de colunas antes da concatenação.
            # Preenche com NaN se o DF tiver menos colunas.
            df_padronizado = df.reindex(columns=range(max_cols_global))

            # Se for o primeiro DataFrame, ele contém o cabeçalho principal.
            # Adicionamos ele diretamente.
            if not primeiro_df_processado:
                df_final_lista.append(df_padronizado)
                primeiro_df_processado = True
            else:
                # Para os DataFrames subsequentes, como não há cabeçalho repetido,
                # apenas adicionamos o DataFrame padronizado diretamente.
                df_final_lista.append(df_padronizado)


        # Concatena todos os DataFrames padronizados
        df_total = pd.concat(df_final_lista, ignore_index=True)

        # Salva o Excel
        df_total.to_excel(excel_path, index=False)
        print(f"✅ Arquivo salvo em: {excel_path}")

    except Exception as e:
        print(f"❌ Erro ao processar {pdf_path}: {e}")
        # Opcional: Salvar o DataFrame parcial para depuração
        # if 'df_total' in locals():
        #     df_total.to_excel("erro_parcial.xlsx", index=False)
        #     print("Arquivo parcial salvo para depuração: erro_parcial.xlsx")

# 👇 Caminhos dos arquivos PDF (você pode usar um loop para vários arquivos)
entrada_pdf = "teste700pag.pdf"
saida_excel = "saida_tabela.xlsx"

processar_pdf_para_excel(entrada_pdf, saida_excel)
