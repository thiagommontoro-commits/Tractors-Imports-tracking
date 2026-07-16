# 📊 Dashboard de Importações de Tratores (Comex Stat)

Este projeto automatiza a extração de dados de importação de tratores da API do Comex Stat e gera um dashboard HTML interativo para análise.

 <!-- Troque este link por um print do seu dashboard -->

---

## ✨ Funcionalidades

- **Extração Automatizada:** Busca dados anuais diretamente da API do governo.
- **Tratamento de Erros:** Lógica de retentativa robusta para lidar com limites de requisição da API (Erro 429).
- **Enriquecimento de Dados:** Criação de uma coluna `HP Bucket` baseada no código NCM para facilitar a análise de segmentos.
- **Dashboard Interativo:** Geração de um arquivo `dashboard_importacoes.html` com gráficos dinâmicos usando Plotly.
- **Automação com GitHub Actions:** O dashboard é atualizado automaticamente a cada alteração no código.

---

## 🤖 Automação com GitHub Actions

Este repositório está configurado para se atualizar sozinho. O processo é o seguinte:

1.  **Gatilhos (Triggers):** Uma automação (workflow) é iniciada por um dos seguintes eventos:
    - **Agendamento:** Toda segunda-feira, às 08:00 (horário UTC).
    - **Push:** Sempre que uma nova alteração é enviada para a branch `main`.
    - **Manual:** Pode ser acionada a qualquer momento pela aba "Actions" do GitHub.
2.  **Execução:** Um servidor do GitHub executa o script `Deixa.py`.
3.  **Geração de Arquivos:** O script extrai os dados mais recentes da API, salva o arquivo `Imports database_ANO-MES-DIA.xlsx` e gera o `dashboard_importacoes.html`.
4.  **Commit:** Se houver qualquer alteração nos dados, o robô do GitHub salva (faz um "commit") os arquivos atualizados (`index.html` e o `.xlsx`) de volta no repositório.

Isso garante que o dashboard e a base de dados estejam sempre sincronizados com a última versão do código e com os dados mais recentes da API.

---

## 🚀 Como Usar

### Visualizando o Dashboard Online

Você pode ver o dashboard mais recente diretamente pelo link do GitHub Pages do projeto.

1.  Vá para `Settings` (Configurações) > `Pages`.
2.  Em `Branch`, selecione `main` e a pasta `/ (root)`.
3.  Clique em `Save`.
4.  Após alguns minutos, seu dashboard estará disponível no link fornecido (algo como `https://SEU-USUARIO.github.io/ComexStat/dashboard_importacoes.html`).
    - **Nota:** Com a alteração para `index.html`, o link principal `https://SEU-USUARIO.github.io/ComexStat/` já irá carregar o dashboard.

### Rodando o Projeto Localmente

1.  **Clone o repositório:**
    ```bash
    git clone https://github.com/SEU-USUARIO/ComexStat.git
    cd ComexStat
    ```

2.  **Crie e ative um ambiente virtual:**
    ```bash
    # Windows
    python -m venv venv
    .\venv\Scripts\activate
    ```

3.  **Instale as dependências:**
    ```bash
    pip install -r requirements.txt
    ```

4.  **Execute o script principal:**
    ```bash
    python Deixa.py
    ```

5.  **Veja o resultado:**
    -   O arquivo `Imports database_ANO-MES-DIA.xlsx` será criado/atualizado.
    -   O arquivo `index.html` será gerado. Abra-o em seu navegador para ver os gráficos.

---

## 🛠️ Tecnologias Utilizadas

- **Python**
- **Pandas:** Para manipulação e análise de dados.
- **Requests:** Para fazer as chamadas à API.
- **Plotly:** Para a criação dos gráficos interativos.
- **GitHub Actions:** Para automação e CI/CD.
