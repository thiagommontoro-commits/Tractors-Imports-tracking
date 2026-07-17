import os
import glob
from datetime import datetime
import pandas as pd
from dotenv import load_dotenv

# Importa as funções dos novos módulos
from extracao import extrair_dados_comex
from dashboard import gerar_dashboard

# ==============================================================================
# ORQUESTRADOR PRINCIPAL
# ==============================================================================

def main():
    """
    Orquestra o processo completo: extrai, salva e gera o dashboard.
    """
    # Carrega as variáveis de ambiente do arquivo .env
    load_dotenv()

    # --- PARÂMETROS DE EXTRAÇÃO ---
    codigos_ncm_tratores = [
        "87019100", "87019200", "87019300", "87019410", "87019490", 
        "87019510", "87019590"
    ]
    anos_desejados = list(range(2022, 2027)) # Anos de 2022 a 2026
    df_para_usar = None

    # --- ETAPA 1: EXTRAÇÃO DE DADOS ---
    print("--- ETAPA 1: TENTANDO EXTRAIR DADOS FRESCOS DO COMEX STAT ---")
    df_fresco = None
    try:
        df_fresco = extrair_dados_comex(
            anos=anos_desejados,
            ncm_codes=codigos_ncm_tratores
        )
    except Exception as e:
        print(f"⚠️  Falha na extração de dados do Comex Stat: {e}")

    if df_fresco is not None and not df_fresco.empty:
        print("✅ Dados frescos extraídos com sucesso do Comex Stat.")
        df_para_usar = df_fresco
        hoje = datetime.now().strftime("%Y-%m-%d")
        nome_arquivo_excel = f"Imports database_{hoje}.xlsx"
        print(f"\n💾 Salvando dados frescos em: {nome_arquivo_excel}")
        df_para_usar.to_excel(nome_arquivo_excel, index=False)
    else:
        print("\n⚠️ Falha ao extrair dados online. Tentando usar dados locais (fallback)...")
        search_pattern = os.path.join(".", "Imports database_*.xlsx")
        lista_arquivos = glob.glob(search_pattern)

        if not lista_arquivos:
            print("❌ NENHUM DADO FRESCO OU LOCAL ENCONTRADO. Abortando a geração do dashboard.")
            return

        caminho_arquivo_antigo = max(lista_arquivos, key=os.path.getctime)
        print(f"📊 Usando dados locais do arquivo: {os.path.basename(caminho_arquivo_antigo)}")
        try:
            df_para_usar = pd.read_excel(caminho_arquivo_antigo)
        except Exception as e:
            print(f"❌ ERRO: Falha ao ler o arquivo Excel local. Causa: {e}")
            return

    if df_para_usar is not None and not df_para_usar.empty:
        # A lógica de geração do dashboard agora está em seu próprio módulo
        gerar_dashboard(df_para_usar)
    else:
        print("❌ Nenhum dado válido para gerar o dashboard.")

if __name__ == "__main__":
    main()