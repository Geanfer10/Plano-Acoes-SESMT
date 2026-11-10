# Plano de Ações de Segurança — SESMT · Goiás Verde

Painel executivo do **Calendário de Ações** do SESMT (CIPA, Auditoria Comportamental, Comitê de Segurança, Campanhas e Acidentes).

## Como atualizar
1. Envie a planilha atualizada para a pasta `uploads/` com o nome **`Calendario_de_Acoes.xlsx`** (substituindo a anterior).
2. O GitHub Actions roda `scripts/extract.py`, gera `dados/acoes.json` e o painel atualiza sozinho em 1–2 minutos.

## Regras
- O status é recalculado no navegador com a data de hoje, com a mesma regra da coluna STATUS da planilha:
  concluída (tem término) · em andamento (prevista não venceu) · em atraso (prevista vencida ou sem data).
- Pendentes = em andamento + em atraso.
- A aba **CONCLUÍDOS** entra como histórico (ações já presentes na aba principal são ignoradas).

## Estrutura
- `index.html` — painel (HTML/JS estático, GitHub Pages)
- `scripts/extract.py` — leitura da planilha → `dados/acoes.json`
- `.github/workflows/processar-planilha.yml` — automação
