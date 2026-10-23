# SmartBid CLI v1.0

## Instalação
```
pip install -r requirements.txt
```
Windows: `setx ANTHROPIC_API_KEY "sk-ant-..."` (abrir novo terminal)
Linux/Mac: `export ANTHROPIC_API_KEY="sk-ant-..."`

## Uso
```
python smartbid.py
```
- **0 – SmartSearch**: busca no PNCP por palavras-chave, UF, valor, órgão e modalidade; prioriza federais/estatais; marca novidades; gera CSV; permite baixar o edital e rodar o fluxo completo.
- **1 – SmartBid**: fluxo integrado Insight → Analyse → Compliance → Position (→ Proposal opcional).
- **2 – Módulos**: cada módulo isolado (Insight, Analyse, Compliance, Position, Proposal, Review, Defense, Sizing, Calc, Intel).

Para arquivos, arraste o arquivo para o terminal (PDF, DOCX, ZIP, TXT).

## Pastas
- `docs_produto/` — coloque aqui datasheets/manuais **oficiais** do fabricante (base das comprovações).
- `saida/` — relatórios DOCX (+ .md) no padrão `CLIENTE_PREGAO_ASSUNTO_DDMMAA.docx`, CSVs do SmartSearch, base SmartIntel e editais baixados.

## Configuração (`config.json`, criado na 1ª execução)
`modelo`, `max_tokens`, `produto`, `pasta_produto`, `keywords`, `ufs`, `valor_min`, `limite_caracteres`.

## Limitações desta versão
- Busca apenas no PNCP (sem ComprasNet/portais estaduais separados) e sem filtro CATMAT.
- Prazos em dias úteis não consideram feriados.
- SmartPosition usa somente a base SmartIntel e históricos fornecidos — não inventa preços.
