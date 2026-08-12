# 📊 Dashboard Comex Stat — Importação de Tratores (automático)

Painel de importação brasileira de tratores por país, ano/mês e segmento de
potência (HP), gerado a partir da API oficial do **Comex Stat / MDIC**.

Metodologia de segmentação: **Global Reporting & Analytics — AGCO (Thiago Montoro)**.

O robô roda **toda semana** (segunda-feira), verifica se o Comex Stat publicou
dados novos e, **só quando há atualização**, republica o dashboard.

---

## 🚀 Como configurar (uma vez só)

### 1. Crie um repositório no GitHub
- Vá em **github.com → New repository**
- Dê um nome (ex.: `dashboard-tratores`)
- Marque como **Public** (necessário para o GitHub Pages gratuito)
- Crie o repositório

### 2. Suba estes arquivos para o repositório
Coloque na **raiz** do repositório:
```
📁 (raiz do repo)
 ├─ gerar_dashboard.py        ← seu script (já ajustado para o GitHub)
 ├─ requirements.txt
 ├─ .gitignore
 └─ .github/
     └─ workflows/
         └─ atualizar-dashboard.yml
```
> ✅ O `gerar_dashboard.py` que acompanha este pacote **já está pronto** para o
> GitHub — não precisa editar nada.

Pode subir pelo site (**Add file → Upload files**) ou pelo Git:
```bash
git init
git add .
git commit -m "Primeira versão do dashboard automático"
git branch -M main
git remote add origin https://github.com/SEU_USUARIO/dashboard-tratores.git
git push -u origin main
```

### 3. Rode uma vez para gerar o primeiro dashboard
- Vá na aba **Actions** do repositório
- Clique em **"Atualizar Dashboard Comex Stat"**
- Clique em **"Run workflow"** → **Run workflow**
- Aguarde ~1–2 min (fica verde ✅). Isso cria o `index.html`.

### 4. Ative o GitHub Pages
- Vá em **Settings → Pages**
- Em **"Build and deployment" → Source**, escolha **"Deploy from a branch"**
- Branch: **main** · Pasta: **/ (root)** → **Save**
- Em ~1 min seu painel estará no ar em:
  ```
  https://SEU_USUARIO.github.io/dashboard-tratores/
  ```

Pronto! 🎉 A partir daí ele se atualiza sozinho toda semana.

---

## 🔄 Como funciona a atualização automática

| Etapa | O que acontece |
|---|---|
| ⏰ Agenda | Toda **segunda-feira 06:00 (Brasília)** o robô acorda |
| 🌐 Coleta | Puxa os dados mais recentes do Comex Stat (2022 → último mês publicado) |
| 🔍 Compara | Verifica se os dados mudaram desde a última vez (ignora o horário de geração) |
| 📤 Publica | **Só se houver novidade**, republica o dashboard. Se não mudou, não faz nada |

Assim, quando sair **agosto/2026**, **setembro/2026** etc., ele traz
automaticamente e atualiza o painel — sem você mexer em nada.

### Rodar na hora (sem esperar a segunda)
Aba **Actions → Atualizar Dashboard Comex Stat → Run workflow**.

### Mudar o dia/horário
No arquivo `.github/workflows/atualizar-dashboard.yml`, ajuste a linha:
```yaml
- cron: "0 9 * * 1"
```
- `0 9 * * 1` = segunda 09:00 UTC (atual)
- `0 9 * * *` = todos os dias 09:00 UTC
- `0 12 * * 3` = quarta 12:00 UTC
> Horário é sempre em **UTC** (Brasília = UTC−3).

---

## 🧩 Arquivos do projeto

| Arquivo | Função |
|---|---|
| `gerar_dashboard.py` | Extrai os dados e gera o `index.html` |
| `index.html` | O dashboard publicado (gerado automaticamente — não edite à mão) |
| `requirements.txt` | Bibliotecas Python necessárias |
| `.github/workflows/atualizar-dashboard.yml` | O robô semanal |
| `.gitignore` | Evita versionar os Excel temporários |

---

## 💻 E no seu PC continua funcionando igual

O script tem um detalhe esperto na função `resolver_pasta_destino()`:
- **No seu PC**: salva na sua pasta `AI Projects\ComexStat` como sempre.
- **No GitHub**: detecta a variável `COMEXSTAT_OUTDIR` (definida pelo workflow) e
  salva o `index.html` na raiz do repositório para publicação.

Ou seja, o **mesmo arquivo** serve para os dois usos, sem conflito.

---

## ❓ Dúvidas comuns

- **O dashboard não aparece?** Verifique se o Pages está ativado (passo 4) e se a
  Action rodou com sucesso (aba Actions, ✅ verde).
- **Quero manter o Excel também?** Ele é gerado a cada rodada, mas não é versionado
  (fica no `.gitignore`). Se quiser guardá-lo, remova a linha `*.xlsx` do `.gitignore`.
- **Repositório privado?** O GitHub Pages gratuito exige repo público. Para privado,
  é necessário um plano GitHub Pro/Team.
