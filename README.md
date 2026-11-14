# Integration Google Sheets

**Make.com** · Automação de planilha → e-mail

Novas linhas no Google Sheets passam pelo filtro **Apto para validação** e disparam um e-mail no Gmail.

## Fluxo

1. **Google Sheets** — Watch New Rows  
2. **Filtro** — Apto para validação  
3. **Gmail** — Send an email  

## Blueprint

Arquivo do cenário (export Make):

[`make/Integration-Google-Sheets.blueprint.json`](make/Integration-Google-Sheets.blueprint.json)

### Importar no Make

1. Abra o Make → **Create a new scenario** (ou um cenário vazio)  
2. Menu **⋯** (acima) → **Import blueprint**  
3. Selecione o arquivo `.json`  
4. Reconecte Google Sheets e Gmail na sua conta  
5. Salve e teste com **Run once**
