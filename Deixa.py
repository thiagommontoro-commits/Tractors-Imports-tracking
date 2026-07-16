import pandas as pd
import plotly.graph_objects as go
import os
import glob
from datetime import datetime
import json

TRANSLATIONS = {
    "pt": {
        "dashboard_title": "Dashboard de Importações",
        "dashboard_h1": "Dashboard de Importações de Tratores",
        "dashboard_subtitle": "Análise de dados da API Comex Stat",
        "developed_by": "Desenvolvido por",
        "data_source": "Fonte dos Dados:",
        "last_update": "Última Atualização:",
        "analyzed_period": "Período Analisado:",
        "global_filter_label": "Filtro Global (KPIs, Barras e Tabela):",
        "kpi_total_value": "Valor Total Importado",
        "kpi_total_qty": "Quantidade Total",
        "kpi_origin_countries": "Países de Origem",
        "chart1_title": "Análise de Importação por País de Origem",
        "fob_value": "Valor US$ FOB",
        "cif_value": "Valor US$ CIF",
        "quantity": "Quantidade",
        "total_fob_value": "Valor Total (US$ FOB)",
        "total_cif_value": "Valor Total (US$ CIF)",
        "total_quantity_axis": "Quantidade Total",
        "table_title": "Análise Anual por País",
        "donut_title": "Distribuição por HP Bucket",
        "donut_filter_label": "Cortes Exclusivos da Rosca:",
        "all_years": "Todos os Anos",
        "all_countries": "Todos os Países",
        "annual_cumulative": "Acumulado Anual",
        "cumulative_until_month": "Acumulado até Mês",
        "no_data_period": "Sem dados para o período.",
        "origin_country_header": "País de Origem",
        "hover_qty": "Qtd",
    },
    "en": {
        "dashboard_title": "Imports Dashboard",
        "dashboard_h1": "Tractor Imports Dashboard",
        "dashboard_subtitle": "Comex Stat API Data Analysis",
        "developed_by": "Developed by",
        "data_source": "Data Source:",
        "last_update": "Last Update:",
        "analyzed_period": "Analyzed Period:",
        "global_filter_label": "Global Filter (KPIs, Bars, and Table):",
        "kpi_total_value": "Total Imported Value",
        "kpi_total_qty": "Total Quantity",
        "kpi_origin_countries": "Origin Countries",
        "chart1_title": "Import Analysis by Country of Origin",
        "fob_value": "FOB Value US$",
        "cif_value": "CIF Value US$",
        "quantity": "Quantity",
        "total_fob_value": "Total Value (US$ FOB)",
        "total_cif_value": "Total Value (US$ CIF)",
        "total_quantity_axis": "Total Quantity",
        "table_title": "Annual Analysis by Country",
        "donut_title": "Distribution by HP Bucket",
        "donut_filter_label": "Donut Chart Exclusive Slices:",
        "all_years": "All Years",
        "all_countries": "All Countries",
        "annual_cumulative": "Annual Cumulative",
        "cumulative_until_month": "Cumulative up to Month",
        "no_data_period": "No data for the period.",
        "origin_country_header": "Country of Origin",
        "hover_qty": "Qty",
    }
}

def format_large_number(num):
    return f"{num:,.0f}".replace(",", ".")

def criar_fig_pais(df_agg):
    fig = go.Figure()
    fig.add_trace(go.Bar(x=df_agg['País de Origem'], y=df_agg['Valor US$ FOB'], name='FOB', text=df_agg['Valor US$ FOB'].apply(format_large_number), textposition='outside'))
    fig.add_trace(go.Bar(x=df_agg['País de Origem'], y=df_agg['Valor US$ CIF'], name='CIF', text=df_agg['Valor US$ CIF'].apply(format_large_number), textposition='outside', visible=False))
    fig.add_trace(go.Bar(x=df_agg['País de Origem'], y=df_agg['Quantidade Estatística'], name='Quantidade', text=df_agg['Quantidade Estatística'].apply(format_large_number), textposition='outside', visible=False))
    fig.update_layout(
        title_text=TRANSLATIONS['pt']['chart1_title'], title_x=0.5,
        updatemenus=[dict(active=0, buttons=list([
            dict(label=TRANSLATIONS['pt']['fob_value'], method="update", args=[{"visible": [True, False, False]}, {"yaxis.title": TRANSLATIONS['pt']['total_fob_value']}]),
            dict(label=TRANSLATIONS['pt']['cif_value'], method="update", args=[{"visible": [False, True, False]}, {"yaxis.title": TRANSLATIONS['pt']['total_cif_value']}]),
            dict(label=TRANSLATIONS['pt']['quantity'], method="update", args=[{"visible": [False, False, True]}, {"yaxis.title": TRANSLATIONS['pt']['total_quantity_axis']}])
        ]), direction="down", x=0.05, xanchor="left", y=1.15, yanchor="top")]
    )
    fig.update_traces(textfont_size=12, textangle=0, cliponaxis=False)
    fig.update_yaxes(title_text="Valor Total (US$ FOB)")
    fig.update_xaxes(categoryorder='total descending')
    return fig

def gerar_dashboard(df: pd.DataFrame = None):
    print("🚀 Iniciando a geração do dashboard...")
    project_path = "."

    if df is None:
        search_pattern = os.path.join(project_path, "Imports database_*.xlsx")
        lista_arquivos = glob.glob(search_pattern)
        if not lista_arquivos:
            print(f"❌ ERRO: Nenhum arquivo Excel encontrado.")
            return

        caminho_arquivo = max(lista_arquivos, key=os.path.getctime)
        print(f"📊 Lendo dados do arquivo: {os.path.basename(caminho_arquivo)}")
        try:
            df = pd.read_excel(caminho_arquivo)
        except Exception as e:
            print(f"❌ ERRO: Falha ao ler o arquivo Excel. Causa: {e}")
            return

    # --- LIMPEZA E BLINDAGEM DE DADOS ---
    numeric_cols = ['Valor US$ FOB', 'Ano', 'Mês', 'Quantidade Estatística', 'Valor US$ CIF']
    for col in numeric_cols:
        if col in df.columns:
            df[col] = pd.to_numeric(df[col], errors='coerce')

    df.dropna(subset=['Valor US$ FOB', 'Ano', 'Mês'], inplace=True)
    df['Mês'] = df['Mês'].astype(int)
    df['Ano'] = df['Ano'].astype(int)

    for col in ['Quantidade Estatística', 'Valor US$ CIF']:
        if col in df.columns: df[col] = df[col].fillna(0)

    ano_min = int(df['Ano'].min())
    ano_max = int(df['Ano'].max())

    metric_options = ['Valor US$ FOB', 'Valor US$ CIF', 'Quantidade Estatística']
    hp_order = ['0 - 24HP', '25 - 50HP', '51 - 100HP', '100 - 175HP', '> 176 HP']
    
    anos_list = sorted(df['Ano'].unique())
    paises_list = sorted(df['País de Origem'].unique())

    # --- GRÁFICOS INICIAIS ---
    print("   -> Gerando Gráfico de Barras Iniciais...")
    initial_country_data = df.groupby('País de Origem')[metric_options].sum().reset_index()
    fig1 = criar_fig_pais(initial_country_data)

    # --- CUBO DE DADOS MASTER ---
    print("   -> Construindo cubo multidimensional de meses, anos e países...")
    master_data = {}
    global_filter_options = [('Total Anual', 'total')] + [(f'Acumulado até Mês {i}', i) for i in range(1, 13)]

    for label, month_key in global_filter_options:
        key = str(month_key)
        
        if month_key == 'total': df_periodo = df
        else: df_periodo = df[df['Mês'] <= month_key]

        kpi_fob = df_periodo['Valor US$ FOB'].sum()
        kpi_qty = df_periodo['Quantidade Estatística'].sum()
        kpi_paises = df_periodo['País de Origem'].nunique()

        country_data = df_periodo.groupby('País de Origem')[metric_options].sum().reset_index()
        if country_data.empty: country_data = pd.DataFrame([{'País de Origem': 'Nenhum dado', 'Valor US$ FOB': 0, 'Valor US$ CIF': 0, 'Quantidade Estatística': 0}])
        fig_pais_periodo = criar_fig_pais(country_data)

        table_data_by_metric = {}
        for metrica in metric_options:
            df_pivot = df_periodo.groupby(['País de Origem', 'Ano'])[metrica].sum().unstack().fillna(0)
            df_pivot = df_pivot.reindex(columns=anos_list, fill_value=0)
            if anos_list: df_pivot = df_pivot.sort_values(by=anos_list[-1], ascending=False)
            
            df_pivot_formatted = df_pivot.map(format_large_number)
            cell_values = [df_pivot_formatted.index.tolist()] + [df_pivot_formatted[ano].tolist() for ano in anos_list]
            table_data_by_metric[metrica] = cell_values

        donut_cube = {}
        opcoes_ano = ['Todos'] + [str(a) for a in anos_list]
        opcoes_pais = ['Todos'] + paises_list

        for y in opcoes_ano:
            donut_cube[y] = {}
            df_y = df_periodo if y == 'Todos' else df_periodo[df_periodo['Ano'] == int(y)]
            
            for p in opcoes_pais:
                df_p = df_y if p == 'Todos' else df_y[df_y['País de Origem'] == p]
                
                if df_p.empty:
                    d_fob = [0]*len(hp_order)
                    d_cif = [0]*len(hp_order)
                    d_qty = [0]*len(hp_order)
                else:
                    hp_agg = df_p.groupby('HP Bucket')[metric_options].sum().reindex(hp_order, fill_value=0).reset_index()
                    d_fob = hp_agg['Valor US$ FOB'].tolist()
                    d_cif = hp_agg['Valor US$ CIF'].tolist()
                    d_qty = hp_agg['Quantidade Estatística'].tolist()

                donut_cube[y][p] = {
                    'labels': hp_order,
                    'fob': d_fob,
                    'cif': d_cif,
                    'qty': d_qty,
                }

        master_data[key] = {
            'kpis': {'fob': f"US$ {format_large_number(kpi_fob)}", 'qty': format_large_number(kpi_qty), 'paises': str(kpi_paises)},
            'chart1': json.loads(fig_pais_periodo.to_json()),
            'chart3_table': table_data_by_metric,
            'donut_cube': donut_cube
        }

    # --- MONTAGEM DO HTML ---
    print("   -> Montando o arquivo HTML final...")
    graph1_div = fig1.to_html(full_html=False, include_plotlyjs='cdn', div_id='chart1_div')
    data_geracao = datetime.now().strftime("%d/%m/%Y às %H:%M")

    # Options HTML blocks para facilitar
    meses_options_html = f'<option value="total" data-translate-key="annual_cumulative">{TRANSLATIONS["pt"]["annual_cumulative"]}</option>' + \
                     ''.join([f'<option value="{i}" data-translate-key="cumulative_until_month" data-month-value="{i}">{TRANSLATIONS["pt"]["cumulative_until_month"]} {i}</option>' for i in range(1, 13)])
    anos_options_html = f'<option value="Todos" data-translate-key="all_years">{TRANSLATIONS["pt"]["all_years"]}</option>' + ''.join([f'<option value="{a}">{a}</option>' for a in anos_list])
    paises_options_html = f'<option value="Todos" data-translate-key="all_countries">{TRANSLATIONS["pt"]["all_countries"]}</option>' + ''.join([f'<option value="{p}">{p}</option>' for p in paises_list])

    html_string = f'''
    <!DOCTYPE html>
    <html lang="pt-br">
    <head>
        <meta charset="UTF-8">
        <title data-translate-key="dashboard_title">{TRANSLATIONS["pt"]["dashboard_title"]}</title>
        <link href="https://fonts.googleapis.com/css2?family=Inter:wght@400;500;600;700&display=swap" rel="stylesheet">
        <script src="https://cdn.plot.ly/plotly-latest.min.js"></script>
        <style>
            :root {{ --header-bg: #1e293b; --agco-red: #CC0000; --bg-page: #f4f7f6; --bg-card: #ffffff; --text-main: #2c3e50; --text-muted: #6c757d; --border-light: #e9ecef; }}
            body {{ font-family: 'Inter', sans-serif; background-color: var(--bg-page); margin: 0; padding: 20px; color: var(--text-main); }}
            .dashboard-container {{ background-color: var(--bg-card); border-radius: 12px; box-shadow: 0 8px 30px rgba(0, 0, 0, 0.05); max-width: 1400px; margin: 0 auto; overflow: hidden; }}
            .header {{ display: flex; justify-content: space-between; align-items: center; background-color: var(--header-bg); padding: 20px 30px; border-bottom: 4px solid var(--agco-red); }}
            .header-controls {{ display: flex; align-items: center; gap: 25px; }}
            .lang-switcher {{ display: flex; gap: 5px; }}
            .lang-btn {{ background-color: rgba(255,255,255,0.1); color: #fff; border: 1px solid #fff; padding: 5px 10px; border-radius: 4px; cursor: pointer; font-size: 0.8em; font-weight: 500; transition: background-color 0.2s, color 0.2s; }}
            .lang-btn:hover {{ background-color: rgba(255,255,255,0.2); }}
            .lang-btn.active {{ background-color: #fff; color: var(--header-bg); font-weight: 700; }}
            .title-area h1 {{ margin: 0 0 5px 0; font-size: 1.8em; font-weight: 700; color: #ffffff; }}
            .title-area p {{ color: #94a3b8; margin: 0; font-size: 1em; }}
            .dev-info {{ text-align: right; color: #94a3b8; font-size: 0.9em; }}
            .dev-info strong {{ color: #ffffff; font-size: 1.1em; display: block; margin-top: 4px; }}
            .content-area {{ padding: 25px 30px 30px 30px; }}
            .info-strip {{ display: flex; flex-wrap: wrap; justify-content: space-between; background-color: #f8f9fa; border-left: 5px solid var(--agco-red); padding: 15px 20px; border-radius: 0 8px 8px 0; margin-bottom: 30px; font-size: 0.95em; color: #495057; gap: 15px; }}
            
            .global-filter-container {{ margin-bottom: 30px; background-color: #fff; padding: 15px 20px; border-radius: 8px; border: 1px solid var(--border-light); display: flex; align-items: center; gap: 15px; }}
            .global-filter-container label {{ font-weight: 600; font-size: 1em; color: var(--text-main); }}
            select.modern-select {{ font-family: 'Inter', sans-serif; font-size: 0.95em; padding: 8px 12px; border-radius: 6px; border: 1px solid #ccc; background-color: #fff; cursor: pointer; min-width: 150px; outline: none; transition: border-color 0.2s; }}
            select.modern-select:focus {{ border-color: var(--agco-red); }}
            
            .kpi-grid {{ display: grid; grid-template-columns: repeat(auto-fit, minmax(240px, 1fr)); gap: 20px; margin-bottom: 35px; }}
            .kpi-card {{ background: var(--bg-card); border: 1px solid var(--border-light); border-left: 4px solid var(--agco-red); border-radius: 8px; padding: 15px 20px; text-align: left; transition: box-shadow 0.2s ease; }}
            .kpi-title {{ font-size: 0.8em; text-transform: uppercase; font-weight: 600; color: var(--text-muted); margin-bottom: 5px; }}
            .kpi-value {{ font-size: 1.8em; font-weight: 700; color: var(--text-main); line-height: 1.2; }}
            .chart-container {{ background-color: var(--bg-card); border: 1px solid var(--border-light); border-radius: 10px; padding: 20px; margin-top: 15px; position: relative; }}
            
            .table-header-area {{ display: flex; justify-content: space-between; align-items: center; margin-bottom: 15px; flex-wrap: wrap; gap: 10px; }}
            .table-title {{ font-size: 1.2em; font-weight: 600; margin: 0; color: var(--text-main); }}
            .btn-group {{ display: flex; gap: 5px; flex-wrap: wrap; }}
            .metric-btn {{ padding: 6px 15px; border: 1px solid #ccc; background: #fff; border-radius: 4px; cursor: pointer; font-size: 0.9em; }}
            .metric-btn.active {{ background: var(--header-bg); color: #fff; border-color: var(--header-bg); font-weight: bold; }}
            
            .custom-table {{ width: 100%; border-collapse: collapse; font-size: 14px; text-align: center; }}
            .custom-table th {{ background-color: var(--header-bg); color: #fff; padding: 12px 10px; border: 1px solid #334155; }}
            .custom-table th:first-child {{ text-align: left; }}
            .custom-table td {{ padding: 10px; border-bottom: 1px solid var(--border-light); }}
            .custom-table td:first-child {{ text-align: left; font-weight: 500; }}
            .custom-table tbody tr:nth-child(even) {{ background-color: #f8fafc; }}
            .custom-table tbody tr:hover {{ background-color: #f1f5f9; }}
            
            .filters-inline {{ display: flex; flex-wrap: wrap; gap: 15px; background: #f8f9fa; padding: 15px; border-radius: 8px; justify-content: center; margin-bottom: 15px; align-items: center; border-left: 4px solid var(--agco-red); }}
        </style>
    </head>
    <body>
        <div class="dashboard-container">
            <div class="header">
                <div class="title-area">
                    <h1 data-translate-key="dashboard_h1">{TRANSLATIONS["pt"]["dashboard_h1"]}</h1>
                    <p data-translate-key="dashboard_subtitle">{TRANSLATIONS["pt"]["dashboard_subtitle"]}</p>
                </div>
                <div class="header-controls">
                    <div class="lang-switcher">
                        <button class="lang-btn active" data-lang="pt">PT</button>
                        <button class="lang-btn" data-lang="en">EN</button>
                    </div>
                    <div class="dev-info"><span data-translate-key="developed_by">{TRANSLATIONS["pt"]["developed_by"]}</span><br><strong>Reporting & Analytics AGCO</strong></div>
                </div>
            </div>
            <div class="content-area">
                <div class="info-strip">
                    <div class="info-item"><strong><span data-translate-key="data_source">{TRANSLATIONS["pt"]["data_source"]}</span></strong> API Comex Stat (MDIC)</div>
                    <div class="info-item"><strong><span data-translate-key="last_update">{TRANSLATIONS["pt"]["last_update"]}</span></strong> {data_geracao}</div>
                    <div class="info-item"><strong><span data-translate-key="analyzed_period">{TRANSLATIONS["pt"]["analyzed_period"]}</span></strong> {ano_min} - {ano_max}</div>
                </div>

                <div class="global-filter-container">
                    <label for="global-time-filter" data-translate-key="global_filter_label">{TRANSLATIONS["pt"]["global_filter_label"]}</label>
                    <select id="global-time-filter" class="modern-select">
                        {meses_options_html}
                    </select>
                </div>

                <div class="kpi-grid">
                    <div class="kpi-card"><div class="kpi-title" data-translate-key="kpi_total_value">{TRANSLATIONS["pt"]["kpi_total_value"]}</div><div class="kpi-value" id="kpi-fob-value">US$ 0</div></div>
                    <div class="kpi-card"><div class="kpi-title" data-translate-key="kpi_total_qty">{TRANSLATIONS["pt"]["kpi_total_qty"]}</div><div class="kpi-value" id="kpi-qty-value">0</div></div>
                    <div class="kpi-card"><div class="kpi-title" data-translate-key="kpi_origin_countries">{TRANSLATIONS["pt"]["kpi_origin_countries"]}</div><div class="kpi-value" id="kpi-paises-value">0</div></div>
                </div>

                <div class="chart-container">{graph1_div}</div>
                
                <div class="chart-container">
                    <div class="table-header-area">
                        <h2 class="table-title" data-translate-key="table_title">{TRANSLATIONS["pt"]["table_title"]}</h2>
                        <div class="btn-group" id="table-btn-group">
                            <button class="metric-btn active" data-metric="Valor US$ FOB" data-translate-key="fob_value">{TRANSLATIONS["pt"]["fob_value"]}</button>
                            <button class="metric-btn" data-metric="Valor US$ CIF" data-translate-key="cif_value">{TRANSLATIONS["pt"]["cif_value"]}</button>
                            <button class="metric-btn" data-metric="Quantidade Estatística" data-translate-key="quantity">{TRANSLATIONS["pt"]["quantity"]}</button>
                        </div>
                    </div>
                    <div id="html_table_container"></div>
                </div>

                <div class="chart-container">
                    <div class="table-header-area">
                        <h2 class="table-title" data-translate-key="donut_title">{TRANSLATIONS["pt"]["donut_title"]}</h2>
                        <div class="btn-group" id="donut-metric-group">
                            <button class="metric-btn donut-metric-btn active" data-metric="fob" data-translate-key="fob_value">{TRANSLATIONS["pt"]["fob_value"]}</button>
                            <button class="metric-btn donut-metric-btn" data-metric="cif" data-translate-key="cif_value">{TRANSLATIONS["pt"]["cif_value"]}</button>
                            <button class="metric-btn donut-metric-btn" data-metric="qty" data-translate-key="quantity">{TRANSLATIONS["pt"]["quantity"]}</button>
                        </div>
                    </div>
                    
                    <div class="filters-inline">
                        <label style="font-size: 0.95em; font-weight: 600;" data-translate-key="donut_filter_label">{TRANSLATIONS["pt"]["donut_filter_label"]}</label>
                        <div>
                            <select id="donut-mes" class="modern-select">
                                {meses_options_html}
                            </select>
                        </div>
                        <div>
                            <select id="donut-ano" class="modern-select">
                                {anos_options_html}
                            </select>
                        </div>
                        <div>
                            <select id="donut-pais" class="modern-select">
                                {paises_options_html}
                            </select>
                        </div>
                    </div>
                    
                    <div id="chart_hp_div" style="width: 100%; height: 450px;"></div>
                </div>
            </div>
        </div>

        <script>
            const masterData = {json.dumps(master_data)};
            const anosColumns = {json.dumps([str(ano) for ano in anos_list])};
            const translations = {json.dumps(TRANSLATIONS)};
            let currentLang = 'pt';
            
            // Filtro Global
            const filterSelect = document.getElementById('global-time-filter');
            
            // Filtros Exclusivos da Rosca
            const donutMesSelect = document.getElementById('donut-mes');
            const donutAnoSelect = document.getElementById('donut-ano');
            const donutPaisSelect = document.getElementById('donut-pais');
            
            const chart1Id = 'chart1_div';
            const chartHpId = 'chart_hp_div';
            const tableContainer = document.getElementById('html_table_container');
            
            let currentTableMetric = 'Valor US$ FOB';
            let currentDonutMetric = 'fob';

            // --- LÓGICA DE TRADUÇÃO ---
            function translatePage(lang) {{
                if (!translations[lang]) return;
                currentLang = lang;
                const langDict = translations[lang];

                document.querySelectorAll('[data-translate-key]').forEach(el => {{
                    const key = el.getAttribute('data-translate-key');
                    if (langDict[key]) {{
                        if (el.hasAttribute('data-month-value')) {{
                            el.textContent = `${{langDict[key]}} ${{el.getAttribute('data-month-value')}}`;
                        }} else {{
                            el.textContent = langDict[key];
                        }}
                    }}
                }});

                // Atualiza componentes dinâmicos
                const currentFilterValue = filterSelect.value;
                const currentData = masterData[currentFilterValue];
                if (!currentData) return;

                // 1. Gráfico de Barras (chart1)
                const chart1Layout = JSON.parse(JSON.stringify(currentData.chart1.layout)); // Deep copy
                chart1Layout.title.text = langDict.chart1_title;
                chart1Layout.updatemenus[0].buttons[0].label = langDict.fob_value;
                chart1Layout.updatemenus[0].buttons[0].args[1]['yaxis.title'] = langDict.total_fob_value;
                chart1Layout.updatemenus[0].buttons[1].label = langDict.cif_value;
                chart1Layout.updatemenus[0].buttons[1].args[1]['yaxis.title'] = langDict.total_cif_value;
                chart1Layout.updatemenus[0].buttons[2].label = langDict.quantity;
                chart1Layout.updatemenus[0].buttons[2].args[1]['yaxis.title'] = langDict.total_quantity_axis;
                
                const visibleTraceIndex = currentData.chart1.data.findIndex(trace => trace.visible === true);
                const yaxisTitleKey = visibleTraceIndex === 0 ? 'total_fob_value' : (visibleTraceIndex === 1 ? 'total_cif_value' : 'total_quantity_axis');
                chart1Layout.yaxis.title.text = langDict[yaxisTitleKey] || langDict.total_fob_value;

                Plotly.react(chart1Id, currentData.chart1.data, chart1Layout);

                // 2. Tabela HTML
                renderHtmlTable(currentData);

                // 3. Gráfico de Rosca (Donut)
                updateDonut();
            }}

            // --- LÓGICA DA TABELA HTML ---
            function renderHtmlTable(newData) {{
                const dataBlock = newData.chart3_table[currentTableMetric];
                const countries = dataBlock[0]; 
                
                if (!countries || countries.length === 0) {{
                    tableContainer.innerHTML = `<p style='text-align:center; padding: 20px; color:#6c757d;'>${{translations[currentLang].no_data_period}}</p>`;
                    return;
                }}

                let html = `<table class="custom-table"><thead><tr><th>${{translations[currentLang].origin_country_header}}</th>`;
                anosColumns.forEach(ano => html += `<th>${{ano}}</th>`);
                html += '</tr></thead><tbody>';

                for (let i = 0; i < countries.length; i++) {{
                    html += `<tr><td>${{countries[i]}}</td>`;
                    for (let j = 1; j <= anosColumns.length; j++) {{
                        html += `<td>${{dataBlock[j][i]}}</td>`;
                    }}
                    html += '</tr>';
                }}
                html += '</tbody></table>';
                tableContainer.innerHTML = html;
            }}

            // --- LÓGICA DA ROSCA (Totalmente Independente) ---
            function updateDonut() {{
                const langDict = translations[currentLang];
                // Agora ela usa SEU PRÓPRIO seletor de mês
                const m = donutMesSelect.value;
                const a = donutAnoSelect.value;
                const p = donutPaisSelect.value;
                
                const currentMonthData = masterData[m];
                if(!currentMonthData || !currentMonthData.donut_cube[a] || !currentMonthData.donut_cube[a][p]) return;
                
                const dData = currentMonthData.donut_cube[a][p];

                let activeValues = [];
                let hoverTemplate = "";
                
                if (currentDonutMetric === 'fob') {{
                    activeValues = dData.fob;
                    hoverTemplate = "<b>%{{label}}</b><br>US$ %{{value:,.0f}} <br>(%{{percent}})<extra></extra>";
                }} else if (currentDonutMetric === 'cif') {{
                    activeValues = dData.cif;
                    hoverTemplate = "<b>%{{label}}</b><br>US$ %{{value:,.0f}} <br>(%{{percent}})<extra></extra>";
                }} else {{
                    activeValues = dData.qty;
                    hoverTemplate = `<b>%{{label}}</b><br>${{langDict.hover_qty}}: %{{value:,.0f}} <br>(%{{percent}})<extra></extra>`;
                }}

                const trace = {{
                    labels: dData.labels,
                    values: activeValues,
                    type: 'pie',
                    hole: 0.45,
                    textinfo: 'percent+label',
                    hovertemplate: hoverTemplate,
                    marker: {{
                        colors: ['#2c3e50', '#CC0000', '#18bc9c', '#f39c12', '#8e44ad']
                    }}
                }};

                const layout = {{
                    margin: {{t: 10, b: 20, l: 10, r: 10}},
                    showlegend: true,
                    legend: {{ orientation: "h", yanchor: "bottom", y: -0.2, xanchor: "center", x: 0.5 }}
                }};

                Plotly.react(chartHpId, [trace], layout);
            }}

            // --- EVENTOS DOS BOTÕES DE IDIOMA ---
            document.querySelectorAll('.lang-btn').forEach(btn => {{
                btn.addEventListener('click', function() {{
                    const lang = this.getAttribute('data-lang');
                    document.querySelectorAll('.lang-btn').forEach(b => b.classList.remove('active'));
                    this.classList.add('active');
                    translatePage(lang);
                }});
            }});

            // --- EVENTOS DOS BOTÕES DA TABELA ---
            const tableBtns = document.querySelectorAll('#table-btn-group .metric-btn');
            tableBtns.forEach(btn => {{
                btn.addEventListener('click', function() {{
                    tableBtns.forEach(b => b.classList.remove('active'));
                    this.classList.add('active');
                    currentTableMetric = this.getAttribute('data-metric');
                    if(masterData[filterSelect.value]) renderHtmlTable(masterData[filterSelect.value]);
                }});
            }});

            // --- EVENTOS DOS BOTÕES DA ROSCA ---
            const donutBtns = document.querySelectorAll('.donut-metric-btn');
            donutBtns.forEach(btn => {{
                btn.addEventListener('click', function() {{
                    donutBtns.forEach(b => b.classList.remove('active'));
                    this.classList.add('active');
                    currentDonutMetric = this.getAttribute('data-metric');
                    updateDonut();
                }});
            }});

            // --- EVENTO DO FILTRO GLOBAL (Afeta apenas KPIs, Barras e Tabela) ---
            filterSelect.addEventListener('change', function() {{
                try {{
                    const newData = masterData[this.value];
                    if (!newData) return;

                    renderHtmlTable(newData);
                    document.getElementById('kpi-fob-value').textContent = newData.kpis.fob;
                    document.getElementById('kpi-qty-value').textContent = newData.kpis.qty;
                    document.getElementById('kpi-paises-value').textContent = newData.kpis.paises;
                    
                    // Re-aplicar tradução ao gráfico que muda
                    const langDict = translations[currentLang];
                    const chart1Layout = JSON.parse(JSON.stringify(newData.chart1.layout)); // Deep copy
                    chart1Layout.title.text = langDict.chart1_title;
                    chart1Layout.updatemenus[0].buttons[0].label = langDict.fob_value;
                    chart1Layout.updatemenus[0].buttons[0].args[1]['yaxis.title'] = langDict.total_fob_value;
                    chart1Layout.updatemenus[0].buttons[1].label = langDict.cif_value;
                    chart1Layout.updatemenus[0].buttons[1].args[1]['yaxis.title'] = langDict.total_cif_value;
                    chart1Layout.updatemenus[0].buttons[2].label = langDict.quantity;
                    chart1Layout.updatemenus[0].buttons[2].args[1]['yaxis.title'] = langDict.total_quantity_axis;
                    
                    const visibleTraceIndex = newData.chart1.data.findIndex(trace => trace.visible === true);
                    const yaxisTitleKey = visibleTraceIndex === 0 ? 'total_fob_value' : (visibleTraceIndex === 1 ? 'total_cif_value' : 'total_quantity_axis');
                    chart1Layout.yaxis.title.text = langDict[yaxisTitleKey] || langDict.total_fob_value;
                    Plotly.react(chart1Id, newData.chart1.data, chart1Layout);
                    
                }} catch (e) {{
                    console.error("Erro ao atualizar dashboard:", e);
                }}
            }});

            // --- EVENTOS DOS FILTROS DA ROSCA (MÊS, ANO E PAÍS EXCLUSIVOS) ---
            donutMesSelect.addEventListener('change', updateDonut);
            donutAnoSelect.addEventListener('change', updateDonut);
            donutPaisSelect.addEventListener('change', updateDonut);

            // --- START DO DASHBOARD (CARGA INICIAL) ---
            renderHtmlTable(masterData['total']);
            document.getElementById('kpi-fob-value').textContent = masterData['total'].kpis.fob;
            document.getElementById('kpi-qty-value').textContent = masterData['total'].kpis.qty;
            document.getElementById('kpi-paises-value').textContent = masterData['total'].kpis.paises;
            
            // Sincroniza o dropdown exclusivo de mês com o global inicialmente
            donutMesSelect.value = 'total';
            updateDonut();

        </script>
    </body>
    </html>
    '''

    dashboard_path = os.path.join(project_path, "dashboard_importacoes.html")
    with open(dashboard_path, 'w', encoding='utf-8') as f:
        f.write(html_string)

    print(f"\n🎉 SUCESSO! O dashboard foi gerado e salvo em:\n👉 {dashboard_path}")

if __name__ == "__main__":
    gerar_dashboard()