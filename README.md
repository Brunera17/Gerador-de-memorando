# Gerador-de-memorando

Gera um memorando para entrega de valor a clientes.

O script lê os PDFs de consulta fiscal de uma empresa (federal, estadual e municipal), extrai débitos, parcelamentos, declarações omissas e certidões, e preenche um modelo `.docx` com o memorando pronto. Também exporta um JSON com os dados extraídos.

## Requisitos

- Python 3
- [`pdfplumber`](https://github.com/jsvine/pdfplumber)

```bash
pip install -r requirements.txt
```

## Uso

```bash
python scripts/processar_fiscal.py <pasta_pdfs> <template.docx> <output.docx> [honorarios]
```

Exemplo:

```bash
python scripts/processar_fiscal.py pdfs templates/modelo_memorando.docx output/Memorando.docx
```

A pasta de PDFs pode usar subpastas por âmbito:

```
pdfs/
├── federal/     Situação Fiscal, Parcelamento SN, Parcelamento Simplificado, Regularize (DAU/SISPAR)
├── estadual/    CRDA PGE-SP, CND SEFAZ-SP, Relatório de Pendências (Site do Contribuinte)
└── municipal/   ISS Web / Certidão Positiva Municipal
```

O `<output.docx>` gera também `<output>_dados.json` ao lado.

## Modelos (`templates/`)

| Arquivo | Uso |
| --- | --- |
| `modelo_memorando.docx` | Memorando de consulta fiscal |
| `modelo_projecao.docx` | Projeção de imposto |

Os modelos usam os placeholders abaixo, substituídos pelo script:

`{{NOME_EMPRESA}}`, `{{CNPJ}}`, `{{DATA_CONSULTA}}`, `{{FEDERAL_TEXTO}}`, `{{ESTADUAL_TEXTO}}`, `{{MUNICIPAL_TEXTO}}`, `{{RESUMO}}`, `{{DATA_ATUALIZACAO_VALORES}}`

As imagens do papel timbrado foram removidas dos modelos deste repositório. Para usar a identidade visual do seu escritório, substitua as imagens em `word/media/` do `.docx` (ou abra o modelo no Word e reinsira o cabeçalho) mantendo os placeholders.

## Dados sensíveis

PDFs de clientes, memorandos gerados e JSONs de dados contêm CNPJs, valores e situação fiscal de terceiros e **não devem ser versionados**. O `.gitignore` já bloqueia `pdfs/`, `*.pdf`, `output/`, `Memorando_*.docx` e `*_dados.json`.
