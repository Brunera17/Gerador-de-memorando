# Gerador-de-memorando

Gera um memorando para entrega de valor a clientes.

O script lê os PDFs de consulta fiscal de uma empresa (federal, estadual e municipal), extrai débitos, parcelamentos, declarações omissas e certidões, e preenche um modelo `.docx` com o memorando pronto. Também exporta um JSON com os dados extraídos.

## Requisitos

- Python 3
- [`pdfplumber`](https://github.com/jsvine/pdfplumber)
- [`streamlit`](https://streamlit.io) (só para a interface web)

```bash
pip install -r requirements.txt
```

## Interface web (sem terminal)

Para quem prefere não usar linha de comando: sobe todos os PDFs juntos (o
programa classifica sozinho), confere os valores extraídos na tela, ajusta
os honorários se quiser, e baixa o `.docx` pronto.

```bash
streamlit run app.py
```

Abre em `http://localhost:8501`.

### Deploy compartilhado (Streamlit Community Cloud)

Para todo mundo do escritório acessar por um link só, sem cada um rodar na
própria máquina:

1. Em [share.streamlit.io](https://share.streamlit.io), conecte a conta
   GitHub e escolha este repositório, branch `main`, arquivo principal
   `app.py`.
2. Em **Settings → Secrets** do app (no painel do Streamlit Cloud, não no
   repositório), cole:
   ```toml
   senha_acesso = "escolha-uma-senha-forte"
   ```
   Sem essa chave configurada o app não pede senha — por isso ela é
   obrigatória para o deploy compartilhado (ver `.streamlit/secrets.toml.example`).
3. Em **Settings → Sharing**, deixe o app como privado (restrito aos
   e-mails da equipe) — a senha acima é uma segunda camada, não substitui
   essa configuração.

Como os PDFs de clientes passam a ser processados no servidor do Streamlit
Cloud (não mais só localmente), só suba PDFs por esse link com a senha
combinada com a equipe, e nunca compartilhe o link publicamente.

## Uso (linha de comando)

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
