# Relatorios de Atendimentos

Este projeto segue o passo a passo do guia fornecido em `2. Proposta de Dashboard Complementar.md` e cria uma base pratica para:

- carregar planilhas CSV;
- normalizar os criterios avaliados;
- calcular indicadores consolidados e por colaborador;
- gerar um relatorio em Markdown com tabelas e graficos;
- gerar um relatorio visual em HTML pronto para impressao;
- permitir a conversao do HTML para PDF;
- documentar a proposta de um dashboard complementar.

## Estrutura

- `generate_report.py`: processa o CSV e gera os relatorios final em Markdown e HTML.
- `dashboard_proposta.md`: traduz a proposta do dashboard para um formato executavel de projeto.
- `export_pdf.ps1`: tenta converter o HTML para PDF com Edge ou Chrome em modo headless.
- `requirements.txt`: dependencias Python.
- `data/`: coloque aqui o CSV bruto.
- `output/`: relatorio final e graficos gerados.

## Como usar

1. Instale as dependencias:

```powershell
python -m pip install -r requirements.txt
```

2. Coloque o CSV de origem em `data/`.

3. Execute o gerador:

```powershell
python .\generate_report.py --input .\data\seu_arquivo.csv --titulo "Relatorio Semanal de Atendimentos Online"
```

Se quiser, voce ainda pode informar `--periodo`, mas o script agora tenta usar automaticamente o intervalo encontrado na coluna de data do CSV.

4. Confira os arquivos gerados em `output/`.

## Saidas geradas

- `output/relatorio_atendimentos.md`
- `output/relatorio_atendimentos.html`
- `output/analise_criterios.png`
- `output/desempenho_colaboradores.png`

## Como gerar PDF

### Opcao 1: pelo navegador

1. Abra `output/relatorio_atendimentos.html`.
2. Clique em `Salvar em PDF` ou use a impressao do navegador.
3. Escolha `Salvar como PDF`.

### Opcao 2: por script

Se houver Edge ou Chrome instalado, execute:

```powershell
powershell -ExecutionPolicy Bypass -File .\export_pdf.ps1
```

Se o navegador nao for encontrado, o script informa isso e voce pode usar a opcao manual acima.

## Observacoes

- O script tenta localizar colunas equivalentes mesmo quando os nomes variam um pouco.
- Existe tratamento especial para colaboradores listados em `SPECIAL_RULES`, como no caso descrito no guia.
- Se o CSV vier sem UTF-8, salve novamente nesse formato antes de executar ou adapte o parametro `--encoding`.
