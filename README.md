# 🍽️ Auditoria de Refeições — SISTERMI

Aplicativo para auditoria automática de refeições cruzando:
- **Excel de refeições** (abas: REFEIÇÕES e EXCEÇÕES)
- **PDF do espelho de ponto** (cartões dos colaboradores)

---

## ⚡ Instalação (1 vez só)

### Pré-requisito: Python 3.10+

Abra o terminal (cmd / PowerShell / Terminal) na pasta do app e execute:

```bash
pip install -r requirements.txt
```

---

## ▶️ Como rodar

```bash
streamlit run app.py
```

O app abrirá automaticamente no navegador em `http://localhost:8501`

---

## 📋 Como usar

1. Faça upload do **Excel** (com abas REFEIÇÕES e EXCEÇÕES)
2. Faça upload do **PDF** do espelho de ponto
3. Clique em **▶ Gerar Auditoria**
4. Visualize os resultados na tela
5. Clique em **📥 Baixar Relatório Excel** para salvar o arquivo

---

## 📊 O que o relatório contém

| Aba | Conteúdo |
|-----|----------|
| **RESUMO** | KPIs gerais + tabela consolidada por colaborador |
| **DETALHE** | Todas as refeições linha a linha com resultado colorido |
| **INCONFORMES** | Somente os casos sem justificativa |
| **EXCEÇÕES APLICADAS** | Todos os casos que usaram exceção |

---

## 🔍 Regras de conformidade

| Tipo | Janela de conformidade |
|------|----------------------|
| ALMOÇO | 11:00 — 17:00 (mínimo 1 min de sobreposição com a jornada) |
| JANTA  | 19:00 — 23:59 (mínimo 1 min de sobreposição com a jornada) |

**Exceções** (aba EXCEÇÕES do Excel): substituem o resultado base.  
- `S` → CONFORME (EXCEÇÃO)  
- `N` → INCONFORME (EXCEÇÃO)

---

## ⚠️ Estrutura esperada do Excel

### Aba REFEIÇÕES
- Linha 1: Período (DATA MEDIÇÃO)
- Linha 2: Restaurante
- Linha 3 (cabeçalho): MATRICULA | NOME | TIPO | DIA_1 ... DIA_31 | TOTAL
- Linha 4+: dados

### Aba EXCEÇÕES
Colunas: `MATRICULA | NOME | DATA_INICIO | DATA_FIM | TIPO | MOTIVO | AUTORIZADO (S/N)`

O campo `DATA_FIM` pode conter o período completo no formato:  
`"01/02/2026 ate 13/02/2026"`

---

## 📞 Suporte
Desenvolvido para SISTERMI — Auditoria de Refeições v1.0
