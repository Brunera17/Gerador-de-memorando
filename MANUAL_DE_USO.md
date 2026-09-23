# Manual de Uso — Gerador de Memorando

Guia para quem vai usar a ferramenta no dia a dia (sem precisar entender
nada de programação). Para instalação, deploy e detalhes técnicos, veja o
[README.md](README.md).

## O que essa ferramenta faz

Você sobe os PDFs de consulta fiscal de uma empresa (federal, estadual e
municipal — não precisa separar), e o programa:

1. Identifica sozinho o tipo de cada PDF.
2. Extrai débitos, parcelamentos, declarações em atraso e certidões.
3. Calcula os honorários de regularização automaticamente.
4. Gera o memorando pronto em `.docx`, já preenchido.

Nenhum dado sai da tela para um serviço externo de IA — toda a extração é
feita por regras fixas no próprio programa.

## Como acessar

- **Local (na sua máquina):** abra um terminal na pasta do projeto e rode:
  ```bash
  streamlit run app.py
  ```
  Abre sozinho em `http://localhost:8501`.
- **Pelo link da nuvem (se o escritório configurou o deploy):** acesse o
  link compartilhado pela equipe. Se pedir senha, é a senha de acesso
  combinada — não compartilhe esse link ou a senha com quem for de fora do
  escritório, já que os PDFs enviados por ali contêm dados de clientes.

## Passo a passo

### 1. Escolha o modelo (barra lateral, à esquerda)

- **Com timbrado do escritório** — para memorandos que vão para o cliente.
- **Sem timbrado (modelo genérico)** — modelo de referência do repositório.

Ainda na barra lateral, se quiser usar um valor fixo de honorários em vez
do cálculo automático, preencha o campo **"Honorários (R$)"**. Deixe em
`0` para o programa calcular sozinho a partir das pendências encontradas.

### 2. Suba os PDFs

Arraste ou selecione todos os PDFs da consulta fiscal daquela empresa —
federal, estadual e municipal juntos, na mesma caixa de upload. Não
precisa organizar em pastas nem renomear nada.

**Documentos que o programa reconhece hoje:**

| Âmbito | Documentos |
| --- | --- |
| Federal | Situação Fiscal (Diagnóstico), Parcelamento do Simples Nacional, Parcelamento Simplificado (e-CAC), Regularize (PGFN — "Vamos Negociar"), Relatório Consolidado da Dívida Ativa, SISPAR, Emissão de Documento de Arrecadação (PGFN), Emissão de Parcela MEI, PGMEI (DAS mensal do MEI), Relatório de Análise Prévia (Simples Nacional) |
| Estadual | CRDA PGE-SP, CND/débitos SEFAZ-SP, Relatório de Pendências Fiscais (Site do Contribuinte — GIA/EFD), Débitos de IPVA, Simulação de Parcelamento ICMS |
| Municipal | ISS/Taxas, Certidão Positiva do Mobiliário, Simulação de Débitos, Listagem de Débito do Mobiliário |

Se um PDF não for reconhecido, ele aparece como `desconhecido` na tela
seguinte e **não entra no memorando** — nesse caso, avise quem mantém o
script (pode ser um formato novo que ainda não foi ensinado ao programa).

### 3. Clique em "Processar PDFs"

O programa lê, classifica e extrai os dados de todos os arquivos enviados.

### 4. Confira antes de gerar

- **Empresa / CNPJ**: confirme que foram identificados corretamente. Se
  aparecer "não identificada/não identificado", confira se o PDF da
  Situação Fiscal (ou equivalente) foi enviado.
- **Classificação de cada PDF enviado**: mostra o que cada arquivo virou
  (✅) ou se ficou como `desconhecido` (⚠️).
- **Resumo por âmbito**: um resumo rápido de Federal, Estadual e
  Municipal — parcelamentos encontrados, atrasos, e um alerta em
  vermelho se a opção pelo Simples Nacional for ser indeferida.

Esse é o momento de perceber se algum PDF importante ficou de fora antes
de gerar o documento final.

### 5. Gere e baixe o memorando

Clique em **"⬇️ Baixar memorando (.docx)"**. O arquivo já sai com o nome
`Memorando_NomeDaEmpresa.docx`, pronto para revisão final e envio ao
cliente.

Se quiser conferir os números por trás do texto (para auditoria ou
dúvida sobre algum valor), abra o expansor **"Dados extraídos (JSON —
para conferência técnica)"** no fim da página.

## Honorários

O cálculo automático segue a mesma tabela de preços do escritório (uma
cobrança por item pendente encontrado — declaração em atraso, parcelamento,
retificação de DAS, etc.). Se a lista de pendências for extensa, o resumo
detalha o valor de cada item antes do subtotal. Para usar um valor fechado
em vez disso, preencha o campo de honorários na barra lateral antes de
gerar o memorando.

## Cuidados com dados sensíveis

- Os PDFs enviados contêm CNPJ, valores e situação fiscal de terceiros —
  não são para serem compartilhados fora do escritório.
- Se estiver usando o link da nuvem, o PDF é processado no servidor do
  Streamlit por alguns segundos e descartado — mesmo assim, use apenas o
  link privado combinado com a equipe, nunca um link público.
- O memorando gerado (`.docx`) também tem dados do cliente — trate-o com
  o mesmo cuidado que qualquer outro documento fiscal do escritório.

## Problemas comuns

| Sintoma | O que fazer |
| --- | --- |
| Empresa/CNPJ não identificados | Confira se o PDF da Situação Fiscal foi enviado |
| PDF aparece como `desconhecido` | Formato ainda não suportado — avise quem mantém o script |
| Valor de um débito parece errado | Abra o expansor de dados extraídos (JSON) e confira o valor bruto extraído daquele PDF |
| "Nenhum modelo .docx encontrado" | Falta o arquivo de modelo na pasta do projeto — avise quem mantém o script |
