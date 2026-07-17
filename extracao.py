import pandas as pd
import ssl

# ==============================================================================
# CONFIGURAÇÃO DE SEGURANÇA
# Desativa a verificação de segurança SSL (Evita o erro CERTIFICATE_VERIFY_FAILED)
# Necessário em redes corporativas ou ao acessar sites do governo (.gov.br)
# ==============================================================================
ssl._create_default_https_context = ssl._create_unverified_context

# Mapeamento de prefixos NCM para categorias de HP para melhor manutenibilidade
HP_BUCKET_MAP = {
    '870191': '25 - 50HP',
    '870192': '51 - 100HP',
    '870193': '101 - 175HP',
    '870194': '> 176 HP',
    '870195': '> 176 HP',
}

def assign_hp_bucket(ncm, bucket_map=HP_BUCKET_MAP):
    """Atribui uma categoria de potência (HP) com base no código NCM."""
    ncm_str = str(ncm)
    for prefix, bucket in bucket_map.items():
        if ncm_str.startswith(prefix):
            return bucket
    return '0 - 24HP' # Categoria padrão

def extrair_dados_comex(anos, ncm_codes):
    """
    Busca os dados de importação das Bases Abertas CSV do portal de economia (MDIC).
    Essa abordagem é à prova de bloqueios de API, pois consome arquivos estáticos.
    """
    print(f"🔎 Buscando dados em Bases Abertas para os anos: {anos} e NCMs: {len(ncm_codes)}")
    
    # 1. Carregar a tabela auxiliar de Países para converter códigos numéricos em nomes
    print("   ⏳ Baixando tabela auxiliar de países...")
    try:
        url_pais = "https://balanca.economia.gov.br/balanca/bd/tabelas/PAIS.csv"
        # O MDIC usa separador de ponto e vírgula e encoding latin-1
        df_pais = pd.read_csv(url_pais, sep=';', encoding='latin-1', usecols=['CO_PAIS', 'NO_PAIS'])
    except Exception as e:
        print(f"   ❌ Erro ao baixar tabela de países. Abortando. Erro: {e}")
        return None

    # 2. Baixar e filtrar os dados anuais
    df_final_list = []
    
    # Colunas que importam para o dashboard (reduz drasticamente o uso de memória RAM)
    colunas_usadas = ['CO_ANO', 'CO_MES', 'CO_NCM', 'CO_PAIS', 'VL_FOB', 'VL_FRETE', 'VL_SEGURO', 'QT_ESTAT']
    
    for ano in anos:
        print(f"   ⏳ Processando dados do ano {ano} (pode levar alguns minutos)...")
        url_ano = f"https://balanca.economia.gov.br/balanca/bd/comexstat-bd/ncm/IMP_{ano}.csv"
        
        try:
            # ESTRATÉGIA DE CHUNKING: Lê o CSV em blocos para não esgotar a memória.
            # Isso é crucial para ambientes como o GitHub Actions.
            chunk_iterator = pd.read_csv(
                url_ano, 
                sep=';', 
                usecols=colunas_usadas, 
                dtype={'CO_NCM': str}, # Garante que o NCM seja tratado como texto
                chunksize=500000  # Processa o arquivo em blocos de 500.000 linhas
            )
            
            chunks_filtrados_ano = []
            for chunk in chunk_iterator:
                # Filtra o bloco atual para os NCMs de interesse
                chunk_filtrado = chunk[chunk['CO_NCM'].isin(ncm_codes)]
                if not chunk_filtrado.empty:
                    chunks_filtrados_ano.append(chunk_filtrado)
            
            df_filtrado = pd.concat(chunks_filtrados_ano, ignore_index=True) if chunks_filtrados_ano else pd.DataFrame()
            
            df_final_list.append(df_filtrado)
            print(f"      ✅ {ano}: {len(df_filtrado)} registros de tratores encontrados.")
            
        except Exception as e:
            # Anos futuros ou não consolidados podem retornar erro se o arquivo não existir
            print(f"      ⚠️ Não foi possível baixar os dados de {ano}. Pode não estar disponível no portal ainda.")

    if not df_final_list:
        print("   ❌ Nenhum dado foi extraído das bases anuais.")
        return None

    # Combina todos os anos filtrados em um único DataFrame
    df_completo = pd.concat(df_final_list, ignore_index=True)

    if df_completo.empty:
        return None

    # --- ENRIQUECIMENTO E LIMPEZA DOS DADOS ---
    print("✨ Enriquecendo e formatando os dados...")
    
    # 3. Cruzamento com a tabela de países (semelhante ao PROCV do Excel)
    df_completo = df_completo.merge(df_pais, on='CO_PAIS', how='left')

    # 4. Cálculo do Valor CIF (FOB + Frete + Seguro)
    # A base aberta não traz o CIF pronto, então calculamos matematicamente
    df_completo['VL_CIF'] = df_completo['VL_FOB'].fillna(0) + df_completo['VL_FRETE'].fillna(0) + df_completo['VL_SEGURO'].fillna(0)

    # 5. Renomear para o padrão esperado pelo dashboard
    rename_map = {
        'CO_ANO': 'Ano',
        'CO_MES': 'Mês',
        'NO_PAIS': 'País de Origem',
        'CO_NCM': 'NCM',
        'VL_FOB': 'Valor US$ FOB',
        'VL_CIF': 'Valor US$ CIF',
        'QT_ESTAT': 'Quantidade Estatística'
    }
    df_completo = df_completo.rename(columns=rename_map)

    # Preenche países não encontrados com "Desconhecido"
    df_completo['País de Origem'] = df_completo['País de Origem'].fillna("Desconhecido")

    # 6. Adiciona a coluna 'HP Bucket'
    df_completo['HP Bucket'] = df_completo['NCM'].apply(assign_hp_bucket)
    print("   ✅ Coluna 'HP Bucket' e Valor CIF processados.")

    # 7. Garante a ordem final das colunas (removendo colunas temporárias como Seguro e Frete)
    final_cols = [
        'Ano', 'Mês', 'País de Origem', 'NCM', 'HP Bucket', 
        'Valor US$ FOB', 'Valor US$ CIF', 'Quantidade Estatística'
    ]
    
    # Filtra apenas as colunas finais
    df_completo = df_completo[final_cols]
    
    print(f"✅ Extração concluída! {len(df_completo)} registros totais preparados para o dashboard.")
    return df_completo