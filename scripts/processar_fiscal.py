#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
processar_fiscal.py v4.0
Uso: python processar_fiscal.py <pasta_pdfs> <template.docx> <output.docx>

Tipos de PDF suportados (por subpasta):
  federal/   SITUACAO FISCAL, Parcelamento SN, Parcelamento Simplificado, Regularize DAU
  estadual/  CRDA PGE-SP, CND SEFAZ-SP
  municipal/ ISS Web
"""

import sys, os, re, json, zipfile, shutil
from datetime import date, datetime, timedelta


# ══════════════════════════════════════════════════════════
# 1. LEITURA DE PDF
# ══════════════════════════════════════════════════════════

def extrair_texto_pdf(caminho):
    try:
        import pdfplumber
        texto = ""
        with pdfplumber.open(caminho) as pdf:
            for p in pdf.pages:
                texto += (p.extract_text() or "") + "\n"
        return texto.strip()
    except ImportError:
        print("ERRO: pip install pdfplumber", file=sys.stderr); sys.exit(1)
    except Exception as e:
        print(f"AVISO: {caminho}: {e}", file=sys.stderr); return ""


def classificar_pdf(texto):
    t = texto.lower()
    if ("secretaria especial da receita federal" in t
            or "informacoes de apoio para emissao de certidao" in t
            or "informações de apoio para emissão de certidão" in t
            or "diagnostico fiscal na receita federal" in t
            or "diagnóstico fiscal na receita federal" in t):
        return "federal"
    if "vamos negociar" in t and "suas dívidas" in t:
        return "regularize_valores"
    if ("vamos simular" in t and "sispar" in t) or             ("vamos simular" in t and "opções de negociação" in t and              ("previdenciário" in t or "simples nacional" in t or "demais débitos" in t)):
        return "parcelamento_sispar"
    if ("detalhamento de inscricao em divida ativa" in t
            or "detalhamento de inscrição em dívida ativa" in t
            or "relatório de inscrições em dívida ativa da união" in t
            or "relatorio de inscricoes em divida ativa da uniao" in t
            or "relatório consolidado da dívida" in t
            or ("regularize" in t and "valor total consolidado" in t and "principal" in t)
            or ("regularize" in t and "dívida total" in t)
            or ("situação da inscrição" in t and "receita da dívida" in t)
            or ("situacao da inscricao" in t and "receita da divida" in t)):
        return "parcelamento_dau"
    if "relatório de análise prévia" in t or "relatorio de analise previa" in t:
        return "analise_previa_sn"
    if "emissão de documento de arrecadação" in t and "negociações:" in t:
        return "parcelamento_pgdau_prestacoes"
    if "parcelas disponíveis para impressão" in t and "parcela" in t and "valor" in t:
        return "parcelamento_mei_emissao"
    if "pgmei" in t and "programa gerador de das" in t and "microempreendedor individual" in t:
        return "mei_pgmei"
    if "parcelamento do simples nacional" in t and "nome empresarial" in t:
        if "valores dos débitos não são suficientes" in t or "valores dos debitos nao sao suficientes" in t:
            return "parc_sn_indisponivel"
        return "parcelamento_sn"
    if "valor da primeira parcela" in t or "valor das demais parcelas" in t:
        return "parcelamento_sn"
    if "saldo a parcelar" in t and "quantidade de parcelas" in t:
        return "parcelamento_simplificado"
    if "nova negociacao" in t and "divida consolidada" in t:
        return "parcelamento_simplificado"
    if "nova negociação" in t and "dívida consolidada" in t:
        return "parcelamento_simplificado"
    if ("parcelamento simplificado" in t and
            ("dívida consolidada não atinge" in t or "divida consolidada nao atinge" in t
             or "mínimo necessário para esta modalidade" in t)):
        return "parcelamento_simplificado"
    if "procuradoria da divida ativa" in t and "estado de sao paulo" in t and "crda" in t:
        return "estadual_pge"
    if "procuradoria da dívida ativa" in t and "estado de são paulo" in t and "crda" in t:
        return "estadual_pge"
    # Site do Contribuinte DEVE vir antes da SEFAZ — o PDF tem rodapé da SEFAZ
    if ("relatório de pendências fiscais" in t
            or "relatorio de pendencias fiscais" in t
            or ("gia/efd" in t and "omisso de declaração" in t)
            or ("gia/efd" in t and "omisso de declaracao" in t)
            or ("icms pendência" in t and "inscrição estadual" in t)):
        return "estadual_site_contribuinte"
    if "débitos de ipva" in t or "debitos de ipva" in t:
        return "estadual_ipva"
    if ("secretaria da fazenda e planejamento do estado de sao paulo" in t
            or "secretaria da fazenda e planejamento do estado de são paulo" in t
            or "debitos tributarios nao inscritos" in t
            or "débitos tributários não inscritos" in t
            or "pfe.fazenda.sp.gov.br" in t):
        return "estadual_sefaz"
    if ("adesão ao parcelamento icms" in t or "adesao ao parcelamento icms" in t
            or ("simulação do parcelamento" in t and "site do contribuinte" in t)
            or ("simulacao do parcelamento" in t and "site do contribuinte" in t)):
        return "estadual_icms_parcelamento"
    if ("iss/taxas" in t or "taxa licenca p/ funcionamento" in t
            or "taxa licença p/ funcionamento" in t
            or "consulta de debitos" in t or "consulta de débitos" in t
            or "certidão positiva do mobiliário" in t
            or "certidao positiva do mobiliario" in t
            or ("prefeitura municipal" in t and "composição da certidão" in t)
            or ("prefeitura municipal" in t and "composicao da certidao" in t)):
        return "municipal"
    return "desconhecido"


# ══════════════════════════════════════════════════════════
# 2. EXTRATORES
# ══════════════════════════════════════════════════════════

def extrair_dados_empresa(textos):
    nome, cnpj = "", ""
    ordem = ["federal","estadual_pge","estadual_sefaz","parcelamento_sn","parcelamento_simplificado","parcelamento_dau","municipal"]
    lista = [textos[k] for k in ordem if k in textos]
    lista += [v for k,v in textos.items() if k not in ordem]
    for texto in lista:
        if not cnpj:
            m = re.search(r'CNPJ[:\s]+(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})', texto, re.IGNORECASE)
            if m: cnpj = m.group(1)
        if not nome:
            # Padrão com CNPJ completo: "CNPJ: XX.XXX.XXX/XXXX-XX - NOME"
            # Padrão com CNPJ base: "CNPJ: XX.XXX.XXX - XX.XXX.XXX NOME" (raiz repetida)
            m2 = re.search(r'CNPJ[:\s]+[\d./\-]+\s*[-–]\s*([\w][^\n]{3,70})', texto, re.IGNORECASE)
            if m2:
                c = m2.group(1).strip()
                # Remover CNPJ base repetido no início (ex: '12.345.678 EMPRESA EXEMPLO' → 'EMPRESA EXEMPLO')
                c = re.sub(r'^\d{2}\.\d{3}\.\d{3}\s+', '', c).strip()
                # Remover CPF numérico colado ao nome
                c = re.sub(r'\s+\d{8,}$', '', c).strip()
                lixo = ["sair","localizar","acesso","mensagens","perfil","dados do","procurador"]
                if not any(l in c.lower() for l in lixo) and len(c) > 3: nome = c
        if not nome:
            m3 = re.search(r'Nome Empresarial[:\s]+([A-ZÁÉÍÓÚÂÊÎÔÛÃÕÀÇ][^\n]{3,70})', texto, re.IGNORECASE)
            if m3: nome = m3.group(1).strip()
        if nome and cnpj: break
    return nome, cnpj


def extrair_federal(texto):
    r = {"status":"NEGATIVA","certidao_numero":"","certidao_validade":"","certidao_tipo":"",
         "tem_debitos_exigiveis":False,"debitos":[],"exigibilidade_suspensa":[],
         "parcelamento_siefpar":None,"parcelamento_ativo":None,"inscricoes_sida":[]}
    if not texto: return r
    m = re.search(r'Certid[aã]o Positiva com Efeitos de Negativa[:\s]+([A-Z0-9.]+)', texto, re.IGNORECASE)
    if m:
        r["certidao_numero"] = m.group(1).strip()
        r["status"] = "POSITIVA_COM_EFEITOS_NEGATIVA"
        r["certidao_tipo"] = "Certidão Positiva com Efeitos de Negativa"
    else:
        # Baseado em um único exemplo real — layout pode variar em outros casos
        m = re.search(r'Certid[aã]o Negativa[:\s]+([A-Z0-9.]+)', texto, re.IGNORECASE)
        if m:
            r["certidao_numero"] = m.group(1).strip()
            r["certidao_tipo"] = "Certidão Negativa"
    m = re.search(r'Data de Validade[:\s]+(\d{2}/\d{2}/\d{4})', texto, re.IGNORECASE)
    if m: r["certidao_validade"] = m.group(1)

    # Tabela SIEF tem colunas:
    # Receita | PA/Exerc | Dt.Vcto | Vl.Original | Sdo.Devedor | Multa | Juros | Sdo.Dev.Cons. | Situacao
    # Usamos Sdo.Dev.Cons. (8a coluna) pois ja inclui multa e juros atualizados
    padrao = re.compile(
        r'^(.+?)\s+(\d{2}/\d{4}|\d{4}|\d+[°º]\s*TRIM/\d{4})\s+(\d{2}/\d{2}/\d{4})\s+'  # Receita PA Vcto
        r'([\d.,]+)\s+([\d.,]+)\s+'                                       # Vl.Orig Sdo.Dev
        r'([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)',                            # Multa Juros Sdo.Dev.Cons
        re.MULTILINE)
    ignorar = {"receita","pa/exerc","dt. vcto","vl. original","sdo. devedor",
               "multa","juros","situação","situacao","sdo. dev. cons."}

    def normalizar_receita(rec):
        """Agrupa receitas do mesmo tipo. Ex: 1082-01 CP-SEGUR + 1082-21 CP-SEGUR -> CP-SEGUR.
           Remove sufixo de trimestre: "IRPJ 2º" → "IRPJ"
        """
        rec = rec.strip()
        # Padrão: "XXXX-XX - NOME" ou "XXXX-XX - NOME."
        m = re.match(r'\d{4}-\d{2}\s*-\s*(.+)', rec)
        if m:
            nome = m.group(1).strip().rstrip(".")
            # Remover sufixo de trimestre: "IRPJ 2º" → "IRPJ", "CSLL 4º" → "CSLL"
            nome = re.sub(r'\s+\d+[°º]$', '', nome).strip()
            return nome
        # Sem código: remover sufixo trimestre
        rec = re.sub(r'\s+\d+[°º]$', '', rec).strip()
        return rec.rstrip(".")

    # Detectar parcelamento com exigibilidade suspensa (PARCSN/PARCMEI)
    # Esses débitos estão suspensos pois estão em parcelamento ativo
    idx_parcsn = max(
        texto.find("Parcelamento com Exigibilidade Suspensa"),
        texto.find("parcelamento com exigibilidade suspensa")
    )
    if idx_parcsn > 0:
        # Capturar tipo do parcelamento (ex: "SIMPLES NACIONAL - EM PARCELAMENTO")
        m_tipo = re.search(r'(SIMPLES NACIONAL|MEI)[^\n]*EM PARCELAMENTO', texto, re.IGNORECASE)
        r["parcelamento_ativo"] = m_tipo.group(0).strip() if m_tipo else "EM PARCELAMENTO"
    else:
        r["parcelamento_ativo"] = None

    # Separar seção de débitos exigíveis da de exigibilidade suspensa (SIEF)
    # A seção PARCSN vem ANTES da seção "Pendência - Débito (SIEF)"
    idx_sief = max(texto.find("Pendência - Débito (SIEF)"), texto.find("pendência - débito (sief)"))
    idx_es = max(texto.find("Débito com Exigibilidade Suspensa"), texto.find("débito com exigibilidade suspensa"))

    # t_exig = apenas a seção Pendência - Débito (SIEF), excluindo exigibilidade suspensa
    if idx_sief > 0 and idx_es > 0 and idx_es > idx_sief:
        t_exig = texto[idx_sief:idx_es]
        t_susp = texto[idx_es:]
    elif idx_sief > 0:
        t_exig = texto[idx_sief:]
        t_susp = ""
    else:
        t_exig = texto
        t_susp = ""

    def parse(trecho):
        items, vistos = [], set()
        # Pré-processar: juntar "4°\nTRIM/AAAA" em uma linha só
        trecho = re.sub(r'(\d+[°º])\s*\n\s*(TRIM/\d{4})', r'\1 \2', trecho)
        # Remover linhas órfãs de TRIM que ficaram sozinhas
        trecho = re.sub(r'\nTRIM/\d{4}\n', '\n', trecho)
        # Remover linhas de "Notificação de lançamento"
        trecho = re.sub(r'Notifica[çc][aã]o de lan[çc]amento[^\n]+\n', '', trecho)
        for m in padrao.finditer(trecho):
            rec_raw = m.group(1).strip()
            if rec_raw.lower() in ignorar or len(rec_raw) < 3: continue
            if not re.search(r'[A-Z0-9]', rec_raw): continue
            rec  = normalizar_receita(rec_raw)
            per  = m.group(2)
            venc = m.group(3)
            orig  = float(m.group(4).replace(".","").replace(",","."))
            try:
                saldo_cons = float(m.group(8).replace(".","").replace(",","."))
            except (IndexError, AttributeError):
                saldo_cons = float(m.group(5).replace(".","").replace(",","."))
            chave = f"{rec}_{per}"
            if chave not in vistos and saldo_cons > 0:
                vistos.add(chave)
                items.append({"receita":rec,"receita_original":rec_raw,
                              "periodo":per,"vencimento":venc,
                              "valor_original":orig,"saldo_devedor":saldo_cons})

        # Segunda passagem: linhas com PA não padrão (ex: "4º", trimestre)
        # Padrão alternativo: receita + data_vencimento + valores (sem PA)
        padrao_alt = re.compile(
            r'^(.+?)\s+(\d{2}/\d{2}/\d{4})\s+'
            r'([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)',
            re.MULTILINE)
        for m in padrao_alt.finditer(trecho):
            rec_raw = m.group(1).strip()
            if rec_raw.lower() in ignorar or len(rec_raw) < 3: continue
            if not re.search(r'[A-Z0-9]', rec_raw): continue
            # Verificar se já foi capturado (evitar duplicatas)
            rec = normalizar_receita(rec_raw)
            venc = m.group(2)
            saldo_cons = float(m.group(7).replace(".","").replace(",","."))
            # Usar vencimento como chave de deduplicação
            # Só adicionar se não foi capturado pela primeira passagem
            # Deduplicar por valor consolidado + vencimento (chave única de negócio)
            ja_tem = any(abs(i["saldo_devedor"] - saldo_cons) < 0.01 and i["vencimento"] == venc
                         for i in items)
            if not ja_tem and saldo_cons > 0:
                vistos.add(chave)
                orig = float(m.group(3).replace(".","").replace(",","."))
                items.append({"receita":rec,"receita_original":rec_raw,
                              "periodo":venc[3:5]+"/"+venc[6:],"vencimento":venc,
                              "valor_original":orig,"saldo_devedor":saldo_cons})
        return items

    r["debitos"] = parse(t_exig)
    if r["debitos"]: r["tem_debitos_exigiveis"] = True
    if t_susp: r["exigibilidade_suspensa"] = parse(t_susp)

    m_parc = re.search(
        r'Parcelamento[:\s]+([\d./-]+)\s+Parcelas em Atraso[:\s]+(\d+)\s+Valor em Atraso[:\s]+([\d.,]+)',
        texto, re.IGNORECASE)
    if m_parc:
        r["parcelamento_siefpar"] = {
            "numero": m_parc.group(1).strip(),
            "parcelas_em_atraso": int(m_parc.group(2)),
            "valor_em_atraso": float(m_parc.group(3).replace(".","").replace(",","."))
        }

    # Extrair inscrições SIDA direto do SITUAÇÃO FISCAL
    idx_sida = max(texto.find('Pendência - Inscrição (SIDA)'),
                   texto.find('pendência - inscrição (sida)'))
    if idx_sida > 0:
        trecho_sida = texto[idx_sida:]
        # Juntar linhas quebradas: '1537-- SIMP NAC -\nMEI' → '1537-- SIMP NAC - MEI'
        trecho_sida = re.sub(r'(-\s*)\n\s*([A-Z]{2,})', r'\1 \2', trecho_sida)
        trecho_sida = re.sub(r'(\w+\.)\n([A-Z]{2,})', r'\1 \2', trecho_sida)
        padrao_insc = re.compile(
            r'(\d{2}\.\d\.\d{2}\.\d+[-\d]*)\s+'
            r'(.+?)(?=\s+\d{2}/\d{2}/\d{4})'
            r'\s+(\d{2}/\d{2}/\d{4})',
            re.DOTALL
        )
        situacoes = re.findall(r'Situação:\s*([^\n]+)', trecho_sida, re.IGNORECASE)
        inscricoes = []
        for idx_i, m in enumerate(padrao_insc.finditer(trecho_sida)):
            receita = m.group(2).strip()[:50]
            if any(x in receita.lower() for x in ['inscrição','receita','ajuizado','processo','tipo','devedor','cnpj']):
                continue
            if len(receita) < 2: continue
            sit = situacoes[idx_i].strip() if idx_i < len(situacoes) else 'ATIVA EM COBRANÇA'
            # Normalizar receitas conhecidas
            receita_norm = receita
            if '1537' in receita and 'SIMP NAC' in receita.upper():
                receita_norm = 'SIMPLES NACIONAL MEI'
            elif '1537' in receita:
                receita_norm = 'SIMPLES NACIONAL'
            inscricoes.append({'inscricao':m.group(1).strip(),'receita':receita_norm,'data':m.group(3),'situacao':sit})
        r['inscricoes_sida'] = inscricoes

    return r


def extrair_declaracoes_omissas(texto):
    """Extrai declarações omissas do SITUAÇÃO FISCAL.
    Tipos suportados: DASN SIMEI, DCTF, ECF, EFD-CONTRIBUIÇÕES, DCTFWeb
    """
    r = {
        "encontrado": False,
        "declaracoes": [],  # [{"tipo": str, "periodos": [str], "multa_unitaria": float}]
        "total_multas": 0.0,
        "inapta_omissao": False
    }
    if not texto:
        return r

    t_lower = texto.lower()

    # Verificar se empresa está inapta por omissão
    if "omissão de declarações" in t_lower or "omissao de declaracoes" in t_lower:
        r["inapta_omissao"] = True

    # ── DASN SIMEI ────────────────────────────────────────────────
    # DASN SIMEI: "(Ano-Calendario) 2023 2024"
    m = re.search(
        r'Omiss.o de DASN SIMEI[^\n]*\n\(Ano-Calend.rio\)\s*(.+?)(?:\n|$)',
        texto, re.IGNORECASE
    )
    if m:
        anos_raw = m.group(1).strip()
        anos = re.findall(r'\d{4}', anos_raw)
        if anos:
            r["declaracoes"].append({
                "tipo": "DASN SIMEI",
                "periodos": anos,
                "descricao": f"Referente aos anos: {_formatar_lista(anos)}.",
                "multa_unitaria": 50.0
            })
            r["encontrado"] = True

    # DCTFWeb
    m_dctfweb = re.search(
        r'Omiss.o de DCTFWeb[^\n]*\n\(Per.odo de Apura..o\)\s*(.+?)(?=\n\n|\n[A-Z*]|\*Aus|$)',
        texto, re.IGNORECASE | re.DOTALL
    )
    if m_dctfweb:
        periodos_raw = m_dctfweb.group(1).strip()
        periodos_fmt = _formatar_periodos_dctf(periodos_raw)
        if periodos_fmt:
            r["declaracoes"].append({"tipo":"DCTFWeb","periodos":[periodos_raw],
                "descricao":f"Referente aos períodos: {periodos_fmt}.","multa_unitaria":250.0})
            r["encontrado"] = True

    # DCTF
    m_dctf = re.search(
        r'Omiss.o de DCTF(?!Web)[^\n]*\n\((?:Ano|Per.odo)[^)]*\)\s*(.+?)(?=\n\n|\n[A-Z*]|$)',
        texto, re.IGNORECASE | re.DOTALL
    )
    if m_dctf:
        anos = re.findall(r"\d{4}", m_dctf.group(1))
        if anos:
            r["declaracoes"].append({"tipo":"DCTF","periodos":anos,
                "descricao":f"Referente aos anos: {_formatar_lista(anos)}.","multa_unitaria":250.0})
            r["encontrado"] = True

    # ECF
    m_ecf = re.search(
        r'Omiss.o de ECF[^\n]*\n\((?:Ano|Per.odo)[^)]*\)\s*(.+?)(?=\n\n|\n[A-Z*]|$)',
        texto, re.IGNORECASE | re.DOTALL
    )
    if m_ecf:
        anos = re.findall(r"\d{4}", m_ecf.group(1))
        if anos:
            r["declaracoes"].append({"tipo":"ECF","periodos":anos,
                "descricao":f"Referente aos anos: {_formatar_lista(anos)}.","multa_unitaria":250.0})
            r["encontrado"] = True

    # EFD-CONTRIBUICOES — alguns relatórios truncam o rótulo para "EFD-CONTRIB",
    # sem o sufixo "UIÇÕES" (limite de coluna do relatório de origem)
    m_efd = re.search(
        r'Omiss.o de EFD.CONTRIB(?:UI..ES)?[^\n]*\n\((?:Ano|Per.odo)[^)]*\)\s*(.+?)(?=\n\n|\n[A-Z*]|$)',
        texto, re.IGNORECASE | re.DOTALL
    )
    if m_efd:
        anos = re.findall(r"\d{4}", m_efd.group(1))
        if anos:
            r["declaracoes"].append({"tipo":"EFD-CONTRIBUIÇÕES","periodos":anos,
                "descricao":f"Referente aos anos: {_formatar_lista(anos)}.","multa_unitaria":250.0})
            r["encontrado"] = True

    # Calcular total de multas
    # DCTFWeb: multa única, não por mês/ano listado no período (ao regularizar a
    # primeira competência em aberto, as demais deixam de ser cobradas).
    total = 0.0
    for decl in r["declaracoes"]:
        if decl["tipo"] == "DCTFWeb":
            total += decl["multa_unitaria"]
        else:
            total += decl["multa_unitaria"] * len(decl["periodos"])
    r["total_multas"] = total

    return r

    return r


def _formatar_lista(items):
    """Formata lista de anos: ['2023', '2024'] → '2023 e 2024'"""
    if not items:
        return ""
    if len(items) == 1:
        return items[0]
    return ", ".join(items[:-1]) + " e " + items[-1]


def _formatar_periodos_dctf(raw):
    """Formata periodos DCTFWeb agrupando por ano."""
    meses_map = {
        "JAN":"jan","FEV":"fev","MAR":"mar","ABR":"abr",
        "MAI":"mai","JUN":"jun","JUL":"jul","AGO":"ago",
        "SET":"set","OUT":"out","NOV":"nov","DEZ":"dez"
    }
    partes = []
    # Buscar padrão "AAAA - MES MES MES..." em todo o texto
    for m_ano in re.finditer(r'(\d{4})\s*-\s*([A-Z\s]+?)(?=\d{4}|\*|$)', raw.upper()):
        ano = m_ano.group(1)
        meses = re.findall(r'\b(?:JAN|FEV|MAR|ABR|MAI|JUN|JUL|AGO|SET|OUT|NOV|DEZ)\b', m_ano.group(2))
        meses_pt = [meses_map.get(m, m.lower()) for m in meses]
        if not meses_pt:
            continue
        if len(meses_pt) == 12:
            partes.append(f"{ano} (jan a dez)")
        else:
            partes.append(f"{ano} ({meses_pt[0]} a {meses_pt[-1]})")
    return " e ".join(partes) if partes else raw[:80]

def extrair_parcelamento_sn(texto):
    r = {"tipo":"simples_nacional","encontrado":False,"indisponivel":False,
         "valor_total":0.0,"num_parcelas":0,"valor_primeira":0.0,"valor_demais":0.0}
    if not texto: return r
    if "valores dos débitos não são suficientes" in texto.lower() or "valores dos debitos nao sao suficientes" in texto.lower():
        r["indisponivel"] = True; r["encontrado"] = True; return r
    m = re.search(r'Valor total consolidado[:\s]+R\$\s*([\d.,]+)', texto, re.IGNORECASE)
    if m:
        r["valor_total"] = float(m.group(1).replace(".","").replace(",",".")); r["encontrado"] = True
    for p in [r'Número de parcelas[^:]*:\s*(\d+)', r'entre 2 e (\d+)']:
        m = re.search(p, texto, re.IGNORECASE)
        if m: r["num_parcelas"] = int(m.group(1)); break
    m = re.search(r'Valor da primeira parcela[:\s]+R\$\s*([\d.,]+)', texto, re.IGNORECASE)
    if m: r["valor_primeira"] = float(m.group(1).replace(".","").replace(",","."))
    m = re.search(r'Valor das demais parcelas[:\s]+R\$\s*([\d.,]+)', texto, re.IGNORECASE)
    if m: r["valor_demais"] = float(m.group(1).replace(".","").replace(",","."))
    return r


def extrair_parcelamento_simplificado(texto):
    r = {"tipo":"simplificado","encontrado":False,"indisponivel":False,
         "principal":0.0,"multa":0.0,
         "juros":0.0,"total":0.0,"num_parcelas":0,"valor_parcela":0.0}
    if not texto: return r
    if ("dívida consolidada não atinge" in texto.lower()
            or "divida consolidada nao atinge" in texto.lower()
            or "mínimo necessário para esta modalidade" in texto.lower()):
        r["indisponivel"] = True
        r["encontrado"]   = True
        return r
    def val(label):
        m = re.search(label + r'[^\n\d]*([\d.,]+)', texto, re.IGNORECASE)
        return float(m.group(1).replace(".","").replace(",",".")) if m else 0.0
    r["principal"]     = val(r'Principal\s*\(BRL\)')
    r["multa"]         = val(r'Multa\s*\(BRL\)')
    r["juros"]         = val(r'Juros\s*\(BRL\)')
    r["total"]         = val(r'Total\s*\(BRL\)')
    r["valor_parcela"] = val(r'Valor das parcelas\s*\(BRL\)')
    m = re.search(r'Quantidade de parcelas[^\d]*(\d+)', texto, re.IGNORECASE)
    if m: r["num_parcelas"] = int(m.group(1))
    if r["total"] > 0 or r["num_parcelas"] > 0: r["encontrado"] = True
    return r


def extrair_parcelamento_dau(texto):
    """Suporta dois formatos Regularize:
    1. Detalhamento individual (uma inscrição com principal/multa/juros)
    2. Relatório Consolidado (múltiplas inscrições com valor consolidado)
    """
    r = {"tipo":"dau","encontrado":False,"inscricao":"","receita":"","data_inscricao":"",
         "situacao":"","principal":0.0,"multa":0.0,"juros":0.0,"encargo":0.0,
         "total":0.0,"protesto":False,"tabelionato":"",
         "inscricoes":[],"total_inscricoes":0}
    if not texto: return r
    r["encontrado"] = True

    # ── Formato 1: Detalhamento individual ──────────────────────────
    m = re.search(r'N[oº°][^\S\n]*inscri..o[:\s]+([\d\s./-]+)', texto, re.IGNORECASE)
    if m: r["inscricao"] = m.group(1).strip().replace(" ","")
    m = re.search(r'Receita da d.vida[^:]*:\s*(.+)', texto, re.IGNORECASE)
    if m: r["receita"] = m.group(1).strip()
    m = re.search(r'Data da inscri..o[^:]*:\s*(\d{2}/\d{2}/\d{4})', texto, re.IGNORECASE)
    if m: r["data_inscricao"] = m.group(1)
    m = re.search(r'Situa..o da inscri..o[^:]*:\s*(.+)', texto, re.IGNORECASE)
    if m: r["situacao"] = m.group(1).strip()
    def ev(label):
        m = re.search(label + r'[^\n]*\n?[^\n]*\n?[^\n]*?R\$\s*([\d.,]+)', texto, re.IGNORECASE)
        return float(m.group(1).replace(".","").replace(",",".")) if m else 0.0
    r["principal"] = ev(r'Principal')
    r["multa"]     = ev(r'Multa')
    r["juros"]     = ev(r'Juros de mora')
    r["encargo"]   = ev(r'Encargo legal')
    r["total"]     = ev(r'Valor total consolidado')
    if "protesto" in texto.lower():
        r["protesto"] = True
        m = re.search(r'TABELIONATO[^.]{0,200}', texto, re.IGNORECASE)
        if m: r["tabelionato"] = m.group(0).strip()[:150]

    # ── Formato 2: Relatório Consolidado (múltiplas inscrições) ─────
    # "Valor total da dívida*: R$ 1.234,56"
    m_total = re.search(r'Valor total da d[íi]vida[^:]*:\s*R\$\s*([\d.,]+)', texto, re.IGNORECASE)
    if not m_total:
        m_total = re.search(r'Valor das inscri[çc][õo]es selecionadas[^:]*:\s*R\$\s*([\d.,]+)', texto, re.IGNORECASE)
    if m_total and r["total"] == 0.0:
        r["total"] = float(m_total.group(1).replace(".","").replace(",","."))

    # Quantidade de inscrições
    m_qtd = re.search(r'Total de inscri[çc][õo]es ativas[^:]*:\s*(\d+)', texto, re.IGNORECASE)
    if not m_qtd:
        m_qtd = re.search(r'Quantidade de inscri[çc][õo]es[^:]*:\s*(\d+)', texto, re.IGNORECASE)
    if m_qtd:
        r["total_inscricoes"] = int(m_qtd.group(1))

    # Linha de cada inscrição: "00 0 00 000000-00  01/01/2025  DEVEDOR  R$ 1.000,00"
    padrao_insc = re.compile(
        r'(\d{2}\s+\d\s+\d{2}\s+[\d-]+)\s+'   # número inscrição
        r'(\d{2}/\d{2}/\d{4})\s+'                  # data
        r'[^\n]+?'                                    # devedor (ignorar)
        r'(ATIVA[^\n]{0,50})\s+'                    # situação
        r'R\$\s*([\d.,]+)',                         # valor
        re.IGNORECASE | re.DOTALL
    )
    for m in padrao_insc.finditer(texto):
        num  = m.group(1).strip().replace(" ",".")
        data = m.group(2)
        sit  = m.group(3).strip()
        val  = float(m.group(4).replace(".","").replace(",","."))
        r["inscricoes"].append({"inscricao":num,"data":data,"situacao":sit,"valor":val})

    # Formato 2b: o PDF do Regularize às vezes quebra a linha da tabela em vários
    # fragmentos (inscrição, situação e valor saem intercalados em linhas soltas),
    # e o padrão acima — que espera tudo numa linha só — não bate com nada.
    # Cada registro ocupa ~5 linhas soltas, e o início da "Situação" costuma vir
    # na linha ANTES do número da inscrição (o resto da coluna quebra depois,
    # junto com data/valor) — por isso cada bloco inclui a linha anterior.
    if not r["inscricoes"]:
        marcador = re.compile(r'\d{2}\s+\d\s+\d{2}\s+[\d-]+')
        linhas = texto.split("\n")
        idx_insc = [i for i, l in enumerate(linhas) if marcador.search(l)]
        vocab_situacao = (r'ATIVA|N[ÃA]O|AJUIZ[ÁA]VEL|AJUIZADA|A|SER|EM|COBRAN[ÇC]A|'
                           r'NEGOCIADAS?|GARANTIDAS?|SUSPENSAS?|EXTINTAS?|NO|SISPAR')
        for j, li in enumerate(idx_insc):
            inicio = max(0, li - 1)
            fim = (idx_insc[j + 1] - 1) if j + 1 < len(idx_insc) else len(linhas)
            bloco = "\n".join(linhas[inicio:fim])
            m_num = marcador.search(bloco)
            m_val = re.search(r'R\$\s*([\d.,]+)', bloco)
            if not m_num or not m_val:
                continue
            m_data = re.search(r'\d{2}/\d{2}/\d{4}', bloco)
            palavras_sit = re.findall(r'\b(?:' + vocab_situacao + r')\b', bloco.upper())
            situacao = " ".join(palavras_sit) if palavras_sit else "ATIVA EM COBRANÇA"
            r["inscricoes"].append({
                "inscricao": m_num.group(0).strip().replace(" ", "."),
                "data": m_data.group(0) if m_data else "",
                "situacao": situacao,
                "valor": float(m_val.group(1).replace(".", "").replace(",", ".")),
            })

    # Receita do consolidado (ex: "Simples Nacional (2)")
    m_rec = re.search(r'(Simples Nacional|COFINS|PIS|IRPJ|CSLL|FGTS)[^\n]*\(\d+\)', texto, re.IGNORECASE)
    if m_rec and not r["receita"]:
        r["receita"] = m_rec.group(1).strip()

    return r


def extrair_pgdau_prestacoes(texto):
    """Extrai o detalhamento de parcelas de um parcelamento de dívida ativa (PGDAU)
    a partir da tela "Emissão de Documento de Arrecadação" do site da PGFN.
    Baseado em um único exemplo real — layout pode variar em outros casos.
    """
    r = {"encontrado": False, "total_parcelas": 0, "valor_consolidado": 0.0,
         "valor_parcela_atual": 0.0, "mes_parcela_atual": "", "parcelas_restantes": 0}
    if not texto:
        return r
    r["encontrado"] = True

    m = re.search(r'Total de Parcelas:\s*(\d+)', texto, re.IGNORECASE)
    if m: r["total_parcelas"] = int(m.group(1))
    # "Valor ... consolidado:" e "Saldo Devedor sem ... Juros:" ficam lado a lado no
    # layout e o pdfplumber intercala rótulo/rótulo/valor/valor em vez de rótulo+valor
    m = re.search(r'Valor\s+Saldo Devedor sem\s*\n\s*([\d.,]+)\s+([\d.,]+)',
                   texto, re.IGNORECASE)
    if m:
        r["valor_consolidado"] = float(m.group(1).replace(".", "").replace(",", "."))
    else:
        m = re.search(r'Valor\s+consolidado:\s*([\d.,]+)', texto, re.IGNORECASE | re.DOTALL)
        if m: r["valor_consolidado"] = float(m.group(1).replace(".", "").replace(",", "."))

    # Linha de cada prestação: "Nr ValorOriginário ValorSdDevedor DataVencPrestação
    # [DataVencDocArrec [NrDocumento]]" — as duas últimas colunas só aparecem quando
    # já existe documento de arrecadação emitido para aquela prestação.
    padrao = re.compile(
        r'^(\d{4})\s+([\d.,]+)\s+([\d.,]+)\s+(\d{2}/\d{2}/\d{4})'
        r'(?:\s+\d{2}/\d{2}/\d{4}(?:\s+\d+)?)?\s*$',
        re.MULTILINE)
    pendentes = []
    for m in padrao.finditer(texto):
        sd_devedor = float(m.group(3).replace(".", "").replace(",", "."))
        if sd_devedor > 0:
            pendentes.append({
                "valor": float(m.group(2).replace(".", "").replace(",", ".")),
                "vencimento": m.group(4),
            })

    # A "parcela atual" é a primeira ainda com saldo devedor — as demais entram
    # como "restantes" (mesmo critério usado nos memorandos já validados)
    if pendentes:
        atual = pendentes[0]
        r["valor_parcela_atual"] = atual["valor"]
        r["mes_parcela_atual"] = atual["vencimento"][3:]  # "MM/AAAA"
        r["parcelas_restantes"] = len(pendentes) - 1

    return r


def extrair_mei_emissao_parcela(texto):
    """Extrai as parcelas em atraso a partir da tela "Emissão de Parcela" do
    eCAC (lista de meses disponíveis para impressão, um valor por mês).
    Baseado em um único exemplo real — layout pode variar em outros casos.
    """
    r = {"encontrado": False, "quantidade": 0, "valor_parcela": 0.0, "total": 0.0}
    if not texto:
        return r
    valores = [float(v.replace(".", "").replace(",", "."))
               for _, v in re.findall(r'(\d{2}/\d{4})\s+R\$\s*([\d.,]+)', texto)]
    if not valores:
        return r
    r["encontrado"] = True
    r["quantidade"] = len(valores)
    r["valor_parcela"] = valores[0]  # valor mensal costuma ser uniforme
    r["total"] = round(sum(valores), 2)
    return r


_MESES_PGMEI = "Janeiro|Fevereiro|Março|Abril|Maio|Junho|Julho|Agosto|Setembro|Outubro|Novembro|Dezembro"

def extrair_pgmei(texto):
    """Extrai os DAS mensais do MEI a partir do PGMEI (Programa Gerador de DAS
    do Microempreendedor Individual). O relatório traz um ano-calendário por
    vez: meses sem apuração ainda calculada aparecem com "-" em todas as
    colunas, os demais já trazem Principal/Multa/Juros/Total — só esses
    entram como débito em atraso. O PGMEI também não libera o ano seguinte
    enquanto a DASN do ano exibido não for entregue — esses anos ficam
    listados à parte, "em aberto", contando um mês por guia até o mês atual.
    Baseado em um único exemplo real — layout pode variar em outros casos.
    """
    r = {"encontrado": False, "ano_calendario": 0, "meses_em_atraso": [], "total_atraso": 0.0, "anos_em_aberto": []}
    if not texto:
        return r
    m_ano = re.search(r'(\d{4})\s+Ok\b', texto)
    if not m_ano:
        return r
    r["encontrado"] = True
    ano = int(m_ano.group(1))
    r["ano_calendario"] = ano

    def parse_valor(s):
        return 0.0 if s == "-" else float(s.replace("R$", "").strip().replace(".", "").replace(",", "."))

    padrao_mes = re.compile(
        rf'({_MESES_PGMEI})/(\d{{4}})\s+(Sim|Não)\s+'
        rf'(-|R\$\s?[\d.,]+)\s+(-|R\$\s?[\d.,]+)\s+(-|R\$\s?[\d.,]+)\s+(-|R\$\s?[\d.,]+)\s+'
        rf'(-|\d{{2}}/\d{{2}}/\d{{4}})\s+(-|\d{{2}}/\d{{2}}/\d{{4}})'
    )
    for m in padrao_mes.finditer(texto):
        mes, ano_mes, _apurado, principal, multa, juros, total, venc, acol = m.groups()
        if total == "-":
            continue
        r["meses_em_atraso"].append({
            "mes": mes, "ano": ano_mes,
            "principal": parse_valor(principal), "multa": parse_valor(multa),
            "juros": parse_valor(juros), "total": parse_valor(total),
            "vencimento": venc, "acolhimento": acol,
        })
    r["total_atraso"] = round(sum(x["total"] for x in r["meses_em_atraso"]), 2)

    # Valor padrão do DAS (só o Principal, sem multa/juros/encargos — esses
    # dependem da data de pagamento, que ainda não existe pra guia futura) —
    # usado como base para estimar o mínimo devido nos anos ainda em aberto.
    valor_padrao = r["meses_em_atraso"][-1]["principal"] if r["meses_em_atraso"] else 0.0

    hoje = date.today()
    for ano_pendente in range(ano + 1, hoje.year + 1):
        qtd_meses = 12 if ano_pendente < hoje.year else hoje.month
        r["anos_em_aberto"].append({
            "ano": ano_pendente,
            "quantidade_guias": qtd_meses,
            "valor_minimo": round(valor_padrao * qtd_meses, 2),
        })

    return r


def extrair_analise_previa_sn(texto):
    """Extrai o resultado do "Relatório de Análise Prévia" do Simples Nacional
    (Lei Complementar nº 123/2006) — avisa se a opção pelo Simples será
    indeferida por causa de débitos em aberto, e a partir de quando.
    Baseado em um único exemplo real — layout pode variar em outros casos.
    """
    r = {"encontrado": False, "indeferida": False, "resultado_texto": "", "data_vigencia": ""}
    if not texto:
        return r
    m = re.search(r'Resultado\s+(.+?)(?=\n|$)', texto, re.IGNORECASE)
    if not m:
        return r
    r["encontrado"] = True
    r["resultado_texto"] = m.group(1).strip()
    if "indeferid" in r["resultado_texto"].lower():
        r["indeferida"] = True
    m = re.search(r'Data de in[íi]cio de vig[êe]ncia\s+(\d{2}/\d{2}/\d{4})', texto, re.IGNORECASE)
    if m:
        r["data_vigencia"] = m.group(1)
    return r


def extrair_regularize_valores(texto):
    """Extrai valores individuais das inscrições do Regularize (Vamos Negociar!)."""
    r = {"encontrado": False, "inscricoes": [], "total": 0.0}
    if not texto or "vamos negociar" not in texto.lower():
        return r

    r["encontrado"] = True

    # Padrão da tabela: "00 0 00 000000-00  1.000,00  01/01/2021  ...  SIM/Não"
    # pdfplumber extrai: "00 0 00 000000-00 1.000,00 SIM" ou sem protesto
    padrao = re.compile(
        r'(\d{2}\s+\d\s+\d{2}\s+[\d-]+)\s+'   # inscrição
        r'([\d.,]+)\s+'                               # valor total
        r'(\d{2}/\d{2}/\d{4}|\d{2}/\d{2}/\n?\d{3}|\d{2}/\d{2}/\d{2,3})',  # data
        re.IGNORECASE
    )

    # Alternativa mais simples: buscar linhas com inscrição + valor
    padrao2 = re.compile(
        r'(\d{2}\s+\d\s+\d{2}\s+[\d-]+)\s+([\d.,]+)\s+(SIM|Não|NAO)?',
        re.IGNORECASE
    )

    # Buscar inscrição + valor + protesto em bloco
    # Padrão: "00 0 00 000000-00 1.000,00 ... SIM" ou "Não" na mesma linha/bloco
    padrao3 = re.compile(
        r'(\d{2}\s+\d\s+\d{2}\s+[\d-]+)\s+([\d.,]+)\s+[\w/\s]*?(SIM|Não|NAO|\b)',
        re.IGNORECASE
    )

    for m in padrao2.finditer(texto):
        num = m.group(1).strip().replace(" ", ".")
        val_str = m.group(2).replace(".", "").replace(",", ".")
        try:
            val = float(val_str)
        except:
            continue
        if val < 1.0:
            continue
        prot_raw = m.group(3) if m.lastindex >= 3 and m.group(3) else "Não"
        # Look ahead in text for SIM/Não near this inscription
        pos = m.end()
        proximo = texto[pos:pos+80]
        if re.search(r'\bSIM\b', proximo, re.IGNORECASE):
            protesto = True
        elif re.search(r'\bNão\b|\bNAO\b', proximo, re.IGNORECASE):
            protesto = False
        else:
            protesto = False
        r["inscricoes"].append({
            "inscricao": num,
            "valor":     val,
            "protesto":  protesto
        })

    # Total
    m_total = re.search(r'Você selecionou R\$\s*([\d.,]+)', texto, re.IGNORECASE)
    if m_total:
        r["total"] = float(m_total.group(1).replace(".", "").replace(",", "."))
    elif r["inscricoes"]:
        r["total"] = sum(i["valor"] for i in r["inscricoes"])

    return r


def extrair_parcelamento_sispar(texto):
    """Extrai opções de parcelamento do SISPAR (simulação PGFN).
    Múltiplos PDFs são concatenados — um por categoria (Previdenciário, SN, Demais).
    """
    if not texto:
        return {"encontrado": False, "opcoes": []}

    result = {"encontrado": True, "opcoes": []}

    # Separar blocos por categoria
    # Cada PDF começa com "Vamos Simular!" e tem uma categoria (Previdenciário, Simples Nacional, Demais Débitos)
    # Separar por "Vamos Simular!" E por "• Opção N" (um PDF pode ter múltiplas opções)
    blocos_raw = re.split(r'Vamos Simular!', texto)
    blocos_opcoes = []
    for br in blocos_raw:
        if not br.strip():
            continue
        # Detectar categoria do bloco
        cat_bloco = ""
        for cat in ["Previdenciário", "Simples Nacional", "Demais Débitos"]:
            if cat.lower() in br.lower():
                cat_bloco = cat
                break
        # Dividir por opções dentro do mesmo bloco
        partes = re.split(r'•\s*Opção\s*\d+', br)
        if len(partes) > 1:
            for parte in partes[1:]:  # ignorar cabeçalho
                blocos_opcoes.append((cat_bloco, parte))
        else:
            blocos_opcoes.append((cat_bloco, br))

    for cat_bloco, bloco in blocos_opcoes:
        if not bloco.strip():
            continue

        opcao = {
            "categoria": cat_bloco,
            "disponivel": True,
            "modalidade": "",
            "num_dividas": 0,
            "total_dividas": 0,
            "total_a_pagar": 0.0,
            "num_prestacoes": 0,
            "valor_prestacao": 0.0,
            "desconto": 0.0
        }

        # Categoria: usar a do bloco pai ou detectar no sub-bloco
        if not opcao["categoria"]:
            for cat in ["Previdenciário", "Simples Nacional", "Demais Débitos"]:
                if cat.lower() in bloco.lower():
                    opcao["categoria"] = cat
                    break

        # Sem opção disponível
        if "não encontramos opção de negociação" in bloco.lower():
            opcao["disponivel"] = False
            result["opcoes"].append(opcao)
            continue

        # Modalidade
        m = re.search(r'Parcelamento Convencional[^\n]*\n(.+)', bloco)
        if m:
            opcao["modalidade"] = m.group(1).strip()[:80]

        # Quantas dívidas negociáveis — o Regularize varia a frase conforme o caso:
        # "N das M dívidas" (parcial), "todas as N dívidas" (100%) ou "a única dívida" (N=1)
        m = re.search(r'Identificamos que (\d+) das (\d+) d[íi]vidas', bloco, re.IGNORECASE)
        if m:
            opcao["num_dividas"]   = int(m.group(1))
            opcao["total_dividas"] = int(m.group(2))
        else:
            m_todas = re.search(r'Identificamos que todas as (\d+) d[íi]vidas', bloco, re.IGNORECASE)
            if m_todas:
                opcao["num_dividas"] = opcao["total_dividas"] = int(m_todas.group(1))
            elif re.search(r'Identificamos que a [uú]nica d[íi]vida', bloco, re.IGNORECASE):
                opcao["num_dividas"] = opcao["total_dividas"] = 1

        # Total a pagar
        m = re.search(r'Total a pagar[^\n]*\n[^\n]*R\$\s*([\d.,]+)', bloco)
        if not m:
            m = re.search(r'Você selecionou R\$\s*([\d.,]+)', bloco)
        if m:
            opcao['total_a_pagar'] = float(m.group(1).replace('.','').replace(',','.'))


        # Número de prestações e valor
        m_prest = re.search(r'Prestações\s+(\d+)', bloco, re.IGNORECASE)
        if m_prest:
            opcao["num_prestacoes"] = int(m_prest.group(1))

        # "Básica  Nx  R$ X"
        m_val = re.search(r'Básica\s+(\d+)x\s+R\$\s*([\d.,]+)', bloco, re.IGNORECASE)
        if m_val:
            opcao["num_prestacoes"]  = int(m_val.group(1))
            opcao["valor_prestacao"] = float(m_val.group(2).replace(".","").replace(",","."))

        # Só incluir opções com valor real (ignorar opções com R$ 0,00)
        if opcao["categoria"] and (opcao["total_a_pagar"] > 0 or not opcao["disponivel"]):
            result["opcoes"].append(opcao)

    return result


def extrair_site_contribuinte(texto):
    """Analisa o Relatório de Pendências Fiscais (Site do Contribuinte SEFAZ-SP)."""
    r = {
        "encontrado": False,
        "cnpj": "",
        "razao_social": "",
        "inscricao_estadual": "",
        "pendencias": [],
        "tipos_sem_debito": [],
        "primeira_omissao": "",
        "ultima_omissao": "",
        "total_competencias": 0
    }
    if not texto:
        return r

    r["encontrado"] = True

    # CNPJ e razão social
    m = re.search(r'(\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2})', texto)
    if m: r["cnpj"] = m.group(1)

    m = re.search(r'Razao Social|Razão Social', texto, re.IGNORECASE)
    if m:
        pos = m.end(); resto = texto[pos:pos+80].strip().split('\n')[0].strip()
        # Remover CPF/números soltos colados ao nome
        resto = re.sub(r'\s+\d{8,}$', '', resto).strip()
        r['razao_social'] = resto

    m = re.search(r'Inscrição Estadual\s+([\d.]+)', texto, re.IGNORECASE)
    if m: r["inscricao_estadual"] = m.group(1).strip()

    # Pendências por tipo (ex: "ICMS Pendência Há Pendências")
    tipos_pendencia = re.findall(r'[A-Z][^\n]+?(?=\s+Há Pendências)', texto)
    tipos_sem = re.findall(r'[A-Z][^\n]+?(?=\s+Não há Débitos)', texto)
    r["tipos_sem_debito"] = [t.strip() for t in tipos_sem]

    # Linhas de GIA/EFD omissa
    # Padrão: CNPJ IE Referencia Documento Pendencia Data
    # GIA/EFD: referência e data podem estar em linhas separadas
    # Ex: "05/2026 GIA/EFD Omisso de Declaração desde o dia\n20/06/2026"
    padrao_gia = re.compile(
        r'(\d{2}/\d{4})\s+GIA/EFD\s+Omisso de Declara[çc][aã]o\s+desde o dia[\s\n]+(\d{2}/\d{2}/\d{4})',
        re.IGNORECASE | re.MULTILINE
    )
    competencias = []
    for m in padrao_gia.finditer(texto):
        competencias.append({
            "referencia": m.group(1),
            "omissa_desde": m.group(2)
        })

    if competencias:
        # Ordenar por referência
        competencias.sort(key=lambda x: (x["referencia"][3:], x["referencia"][:2]))
        r["pendencias"] = competencias
        r["total_competencias"] = len(competencias)
        r["primeira_omissao"] = competencias[0]["referencia"]
        r["ultima_omissao"]   = competencias[-1]["referencia"]
        r["primeira_data"]    = competencias[0]["omissa_desde"]

    return r


def extrair_estadual(texto_pge, texto_sefaz):
    r = {"status":"NEGATIVA","certidaos":[],"tem_debitos":False}
    def dv(emi, dias):
        try:
            dt = datetime.strptime(emi, "%d/%m/%Y")
            return (dt + timedelta(days=dias)).strftime("%d/%m/%Y")
        except: return ""
    def dm(emi, meses):
        try:
            dt = datetime.strptime(emi, "%d/%m/%Y")
            mo = dt.month+meses; yr = dt.year+(mo-1)//12; mo = ((mo-1)%12)+1
            return f"{dt.day:02d}/{mo:02d}/{yr}"
        except: return ""
    if texto_pge:
        mn = re.search(r'Certid[aã]o n[oº]+\s+(\d+)', texto_pge, re.IGNORECASE)
        me = re.search(r'Data e hora da emiss[aã]o\s+(\d{2}/\d{2}/\d{4})', texto_pge, re.IGNORECASE)
        mv = re.search(r'Validade\s+(\d+)', texto_pge, re.IGNORECASE)
        num = mn.group(1) if mn else ""; emi = me.group(1) if me else ""
        dias = int(mv.group(1)) if mv else 30
        neg = "não constam débitos" in texto_pge.lower() or "nao constam debitos" in texto_pge.lower()
        r["certidaos"].append({"orgao":"PGE-SP","numero":num,"emissao":emi,"validade":dv(emi,dias),"resultado":"NEGATIVA" if neg else "POSITIVA"})
        if not neg: r["tem_debitos"] = True; r["status"] = "COM_DEBITOS"
    if texto_sefaz:
        mn = re.search(r'Certid[aã]o n[oº]+\s+([\d-]+)', texto_sefaz, re.IGNORECASE)
        me = re.search(r'Data e hora da emiss[aã]o\s+(\d{2}/\d{2}/\d{4})', texto_sefaz, re.IGNORECASE)
        num = mn.group(1) if mn else ""; emi = me.group(1) if me else ""
        neg = "não constam débitos" in texto_sefaz.lower() or "nao constam debitos" in texto_sefaz.lower()
        r["certidaos"].append({"orgao":"SEFAZ-SP","numero":num,"emissao":emi,"validade":dm(emi,6),"resultado":"NEGATIVA" if neg else "POSITIVA"})
        if not neg: r["tem_debitos"] = True; r["status"] = "COM_DEBITOS"
    return r


def extrair_ipva(texto):
    """Extrai débitos de IPVA do relatório "Débitos de IPVA Vinculados ao Sujeito
    Passivo" (Secretaria da Fazenda e Planejamento do Estado de São Paulo).
    Baseado em um único exemplo real — layout pode variar em outros casos
    (ex.: mais de um veículo).
    """
    r = {"encontrado": False, "veiculos": [], "total": 0.0}
    if not texto:
        return r
    m_placa = re.search(r'\b([A-Z]{3}\d[A-Z0-9]\d{2})\s+(\d+)\s+[\d-]+\s+(\S+)', texto)
    m_valor = re.search(r'(\d{4})\s+R\$\s*([\d.,]+)\s+\w+\s+(\d+)\s+\w+', texto)
    if not (m_placa and m_valor):
        return r
    r["encontrado"] = True
    valor = float(m_valor.group(2).replace(".", "").replace(",", "."))
    r["veiculos"].append({
        "placa": m_placa.group(1),
        "renavam": m_placa.group(2),
        "municipio": m_placa.group(3),
        "exercicio": m_valor.group(1),
        "valor": valor,
    })
    r["total"] = round(sum(v["valor"] for v in r["veiculos"]), 2)
    return r


def extrair_icms_parcelamento(texto):
    """Extrai a simulação de parcelamento de ICMS do Site do Contribuinte (SEFAZ-SP).
    Baseado em um único exemplo real — layout pode variar em outros casos.
    """
    r = {"encontrado": False, "quantidade_dividas": 0, "total": 0.0,
         "num_parcelas": 0, "valor_parcela": 0.0}
    if not texto:
        return r

    # "DÉBITOS SELECIONADOS" — linha com 6 valores em R$ + quantidade de dívidas (inteiro)
    # Principal | Juros Moratórios | Multas | Hon. Advocatícios | Hon. Administrativos | Total | Quantidade
    idx = texto.upper().find("DÉBITOS SELECIONADOS")
    if idx < 0:
        idx = texto.upper().find("DEBITOS SELECIONADOS")
    if idx < 0:
        return r
    m = re.search(
        r'([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+(\d+)',
        texto[idx:])
    if not m:
        return r
    r["encontrado"] = True
    r["total"] = float(m.group(6).replace(".", "").replace(",", "."))
    r["quantidade_dividas"] = int(m.group(7))

    # "RESUMO DO PARCELAMENTO" — linha com os mesmos valores + acréscimo financeiro +
    # quantidade de parcelas do plano simulado, terminando em "Simular"
    m_resumo = re.search(
        r'([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+(\d+)\s+Simular',
        texto)
    if m_resumo:
        r["num_parcelas"] = int(m_resumo.group(8))
        total_plano = float(m_resumo.group(7).replace(".", "").replace(",", "."))
        if r["num_parcelas"] > 0:
            r["valor_parcela"] = round(total_plano / r["num_parcelas"], 2)
        if r["total"] == 0.0:
            r["total"] = total_plano

    return r


def extrair_municipal(texto):
    r = {"status":"NEGATIVA","tem_debitos":False,"debitos":[],"total":0.0,
         "certidao_numero":"","certidao_tipo":"","parcelamento_opcoes":[],
         "debitos_parcelamento":[],"total_parcelamento":0.0}
    if not texto: return r

    # ── Formato 1: ISS Web (EX/AT/AJ) ───────────────────────────────
    m = re.search(r'Total Selecionado\s+([\d.,]+)', texto, re.IGNORECASE)
    if m: r["total"] = float(m.group(1).replace(".","").replace(",","."))
    padrao_iss = re.compile(
        r'\b(EX|AT|AJ|C)\b\s+(\d+)\s+(\d{4})\s+(\d+)\s+[NS]\s+([A-ZÁÉÍÓÚ /]+?)(?=\s+\d{2}/)',
        re.IGNORECASE)
    sm = {"EX":"Dívida do Ano","AT":"Dívida Ativa","AJ":"Dívida Ajuizada","C":"Cartório"}
    for m in padrao_iss.finditer(texto):
        r["debitos"].append({"status":sm.get(m.group(1).upper(),m.group(1)),
                             "numero":m.group(2),"ano":m.group(3),
                             "parcela":m.group(4),"receita":m.group(5).strip(),
                             "valor":0.0,"a_pagar":0.0})

    # ── Formato 2: Certidão Positiva Municipal (Fiorilli/Prefeitura) ─
    m_cert = re.search(r'Certid[aã]o Positiva N[°º]\s*([\d/]+)', texto, re.IGNORECASE)
    if m_cert:
        r["certidao_numero"] = m_cert.group(1)
        r["certidao_tipo"]   = "POSITIVA"

    # Tabela: Divida Parc. Ano Vencimento Receita Valor Correção Multa Juros A Pagar
    # Municipal Fiorilli: Divida Parc Ano Vencimento Receita Valor Correção Multa Juros A Pagar
    # Receita pode ter múltiplas palavras: "TAXA LICENÇA P/ FUNCIONAMENTO"
    # Estratégia: pegar tudo entre a data e o primeiro número de valor
    padrao_cert = re.compile(
        r'(\d+)\s+(\d+)\s+(\d{4})\s+(\d{2}/\d{2}/\d{4})\s+'
        r'([A-ZÁÉÍÓÚÂÊÎÔÛÃÕÀÇ][A-ZÁÉÍÓÚÂÊÎÔÛÃÕÀÇ\s/\-]+?)\s+'
        r'([\d.,]+)\s+[\d.,]+\s+[\d.,]+\s+[\d.,]+\s+([\d.,]+)',
        re.IGNORECASE)
    for m in padrao_cert.finditer(texto):
        val    = float(m.group(6).replace(".","").replace(",","."))
        a_pag  = float(m.group(7).replace(".","").replace(",","."))
        r["debitos"].append({
            "status":  "Certidão Positiva",
            "numero":  m.group(1),
            "ano":     m.group(3),
            "parcela": m.group(2),
            "receita": m.group(5).strip(),
            "valor":   val,
            "a_pagar": a_pag
        })

    # ── Formato 3: Simulação de Débitos (Fiorilli — ex.: Prefeitura de Itaí) ─
    # Vários blocos "Exercício: AAAA - Código da Dívida: NNNNNNN", cada um com
    # linhas "Mod Receita Vencimento Parc Valor Desconto Correção Multa Juros
    # Honorários À pagar Situação" (o cabeçalho junta "Juros"+"Honorários" sem
    # espaço). Baseado em um único exemplo real — layout pode variar.
    if "simulação de débitos" in texto.lower() or "simulacao de debitos" in texto.lower():
        for bloco_ex in re.finditer(
                r'Exerc[íi]cio:\s*(\d{4})\s*-\s*C[óo]digo da D[íi]vida:\s*(\d+).*?'
                r'(?=Exerc[íi]cio:\s*\d{4}\s*-\s*C[óo]digo|Sub\.\s*Total|NP\s+Primeira|$)',
                texto, re.IGNORECASE | re.DOTALL):
            ano, codigo, trecho = bloco_ex.group(1), bloco_ex.group(2), bloco_ex.group(0)
            padrao_sim = re.compile(
                r'(\d+)\s+([A-ZÁÉÍÓÚÂÊÎÔÛÃÕÀÇ][A-ZÁÉÍÓÚÂÊÎÔÛÃÕÀÇ\s/\.\-]+?)\s+'
                r'(\d{2}/\d{2}/\d{4})\s+(\d+)\s+'
                r'([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+'
                r'([A-ZÀ-Ü][A-ZÀ-Ü\s]*?)(?=\n|\d|$)')
            for m in padrao_sim.finditer(trecho):
                r["debitos"].append({
                    "status":  m.group(12).strip(),
                    "numero":  codigo,
                    "ano":     ano,
                    "parcela": m.group(4),
                    "receita": m.group(2).strip(),
                    "valor":   float(m.group(5).replace(".","").replace(",",".")),
                    "a_pagar": float(m.group(11).replace(".","").replace(",",".")),
                })

        # Total geral: última linha "Total: ..." (soma de todos os Exercícios)
        m_tot_sim = re.search(r'(?<!Sub\.\s)Total:\s+[\d.,]+(?:\s+[\d.,]+){5}\s+([\d.,]+)',
                               texto, re.IGNORECASE)
        if m_tot_sim and r["total"] == 0.0:
            r["total"] = float(m_tot_sim.group(1).replace(".","").replace(",","."))

        # Opções de parcelamento: "NP Primeira Demais Desconto" (repetido em colunas),
        # seguido de linhas com grupos de 4 números (NP, Primeira, Demais, Desconto)
        idx_np = texto.find("NP Primeira Demais Desconto")
        if idx_np >= 0:
            idx_fim = texto.find("Fiorilli", idx_np)
            trecho_np = texto[idx_np:idx_fim if idx_fim > 0 else len(texto)]
            trecho_np = re.sub(r'NP Primeira Demais Desconto', '', trecho_np, flags=re.IGNORECASE)
            opcoes = []
            for m in re.finditer(r'(\d+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)', trecho_np):
                opcoes.append({
                    "num_parcelas": int(m.group(1)),
                    "primeira":     float(m.group(2).replace(".","").replace(",",".")),
                    "demais":       float(m.group(3).replace(".","").replace(",",".")),
                    "desconto":     float(m.group(4).replace(".","").replace(",",".")),
                })
            if opcoes:
                r["parcelamento_opcoes"] = opcoes

    # Total da certidão: "Totais: 339,64 ... 409,56"
    m_tot = re.search(r'Totais:\s+[\d.,]+(?:\s+[\d.,]+){3}\s+([\d.,]+)', texto, re.IGNORECASE)
    if m_tot and r["total"] == 0.0:
        r["total"] = float(m_tot.group(1).replace(".","").replace(",","."))

    # ── Formato 4: Listagem de Débito do Mobiliário por Dívida (Fiorilli) ───
    # Usada quando a prefeitura já separou os débitos que entram no parcelamento
    # (por isso o valor daqui tem prioridade sobre o da certidão: é o que
    # efetivamente vai ser parcelado, não só o retrato da dívida no momento).
    # Baseado em um único exemplo real — layout pode variar.
    if "listagem de débito" in texto.lower() or "listagem de debito" in texto.lower():
        padrao_list = re.compile(
            r'(\d{4})\s+(\d+)\s+([A-ZÁÉÍÓÚÂÊÎÔÛÃÕÀÇ][A-ZÁÉÍÓÚÂÊÎÔÛÃÕÀÇ\s/\.\-]+?)\s+'
            r'([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)')
        for m in padrao_list.finditer(texto):
            r["debitos_parcelamento"].append({
                # Esse formato não traz uma coluna de situação — inventar um rótulo
                # (ex.: "A Parcelar") ficava sem sentido quando a própria receita já
                # é "PARCELAMENTO" (um débito que já é parcela de algo em andamento).
                "status":  "",
                "numero":  m.group(2),
                "ano":     m.group(1),
                "parcela": "",
                "receita": m.group(3).strip(),
                "valor":   float(m.group(4).replace(".","").replace(",",".")),
                "a_pagar": float(m.group(10).replace(".","").replace(",",".")),
            })
        m_tot_list = re.search(r'Total Geral:\s*[\d.,]+(?:\s+[\d.,]+){5}\s+([\d.,]+)', texto, re.IGNORECASE)
        if m_tot_list:
            r["total_parcelamento"] = float(m_tot_list.group(1).replace(".","").replace(",","."))

    if r["debitos"] or r["total"] > 0 or r["total_parcelamento"] > 0:
        r["tem_debitos"] = True
        r["status"]      = "COM_DEBITOS"
    return r


# ══════════════════════════════════════════════════════════
# 3. XML RICH FORMAT
# ══════════════════════════════════════════════════════════

def escape_xml(t):
    if not isinstance(t, str): t = str(t)
    return t.replace("&","&amp;").replace("<","&lt;").replace(">","&gt;").replace('"','&quot;')

def fmt(v):
    return f"R$ {v:,.2f}".replace(",","X").replace(".",",").replace("X",".")

def plural(qtd, singular, plural_):
    """Forma singular/plural correta conforme a quantidade — nunca "(ões)"/"(is)"."""
    return singular if qtd == 1 else plural_

def make_rpr(bold=False, italic=False, color=None):
    p = []
    if bold: p.append("<w:b/><w:bCs/>")
    if italic: p.append("<w:i/><w:iCs/>")
    if color: p.append(f'<w:color w:val="{color}"/>')
    return ("<w:rPr>"+"".join(p)+"</w:rPr>") if p else ""

def wr(texto, bold=False, italic=False, color=None):
    rpr = make_rpr(bold=bold, italic=italic, color=color)
    esc = escape_xml(str(texto))
    sp  = ' xml:space="preserve"' if (texto != texto.strip() or texto.startswith(" ") or texto.endswith(" ")) else ""
    return f'<w:r>{rpr}<w:t{sp}>{esc}</w:t></w:r>'

def wp(indent=False, before=0, after=120):
    ind = '<w:ind w:left="360" w:hanging="180"/>' if indent else ""
    return f'<w:p><w:pPr><w:spacing w:before="{before}" w:after="{after}" w:line="276" w:lineRule="auto"/>{ind}</w:pPr>'

def wp_sep():
    return '<w:p><w:pPr><w:spacing w:before="60" w:after="120"/><w:pBdr><w:bottom w:val="single" w:sz="6" w:space="1" w:color="AAAAAA"/></w:pBdr></w:pPr></w:p>'

def wp_bullet(normal="", bold_txt=""):
    return ('<w:p><w:pPr><w:spacing w:before="0" w:after="80" w:line="276" w:lineRule="auto"/><w:ind w:left="360" w:hanging="180"/></w:pPr>'
            + wr("• ", bold=True)
            + (wr(normal) if normal else "")
            + (wr(bold_txt, bold=True) if bold_txt else "")
            + '</w:p>')

def blocos_para_xml(blocos):
    xml = []
    for b in blocos:
        t = b.get("type","")
        if t == "italic":
            xml.append(wp(before=80, after=40) + wr(b["text"], italic=True) + "</w:p>")
        elif t == "bullet":
            xml.append(wp_bullet(b.get("normal",""), b.get("bold","")))
        elif t == "label":
            xml.append(wp(before=120, after=40) + wr(b["text"], bold=True) + "</w:p>")
        elif t == "normal":
            xml.append(wp(before=0, after=60) + wr(b["text"]) + "</w:p>")
        elif t == "alerta":
            xml.append(wp(before=80, after=100) + wr(b["text"], bold=True, color="C00000") + "</w:p>")
        elif t == "normal_bold":
            xml.append(wp(before=0, after=60) + (wr(b["text"]) if b.get("text") else "") + (wr(b["bold"], bold=True) if b.get("bold") else "") + "</w:p>")
        elif t == "date":
            xml.append(wp(before=60, after=100) + wr(b["text"], italic=True, color="595959") + "</w:p>")
        elif t == "separator":
            xml.append(wp_sep())
        elif t == "divida":
            titulo = "Dívida Ativa — "
            if b.get("todas_negociadas"):
                titulo = "Dívida Ativa (já em parcelamento) — "
            xml.append(wp(before=100, after=40) + wr(titulo, bold=True) + wr(b.get("receita",""), bold=True) + "</w:p>")
            # Múltiplas inscrições (relatório consolidado)
            if b.get("inscricoes"):
                for insc in b["inscricoes"]:
                    xml.append(wp_bullet(
                        normal=f'Inscrição nº ',
                        bold_txt=f'{insc["inscricao"]}' +
                                 f' — {insc["situacao"]} — ' +
                                 f'{fmt(insc["valor"]) if isinstance(insc["valor"], float) else insc["valor"]}'
                    ))
                xml.append(wp(before=0, after=40) + wr("Valor total consolidado: ") + wr(b.get("total",""), bold=True) + "</w:p>")
            else:
                # Inscrição individual
                xml.append(wp(before=0, after=40) + wr("Inscrição nº ") + wr(b.get("inscricao",""), bold=True) + wr(f' — inscrita em {b.get("data","")}') + "</w:p>")
                xml.append(wp(before=0, after=40) + wr(f'Principal: {b.get("principal","")} | Multa: {b.get("multa","")} | Juros: {b.get("juros","")} | Encargo: {b.get("encargo","")}') + "</w:p>")
                xml.append(wp(before=0, after=40) + wr("Valor total consolidado: ") + wr(b.get("total",""), bold=True) + wr(" — Situação: ") + wr(b.get("situacao",""), bold=True) + "</w:p>")
            if b.get("todas_negociadas"):
                xml.append(wp(before=0, after=40) + wr(
                    "* Inscrições já negociadas — parcelamento em andamento, não é necessário negociar novamente.",
                    italic=True, color="595959") + "</w:p>")
            elif b.get("situacao_negociada"):
                xml.append(wp(before=0, after=40) + wr(
                    "* Parte das inscrições já está negociada (parcelamento em andamento) — ver situação de cada uma acima.",
                    italic=True, color="595959") + "</w:p>")
            if b.get("protesto"):
                xml.append(wp(before=0, after=40) + wr("⚠ Inscrição enviada para protesto — regularização no tabelionato.", bold=True, color="C00000") + "</w:p>")
        elif t == "parc_sn":
            xml.append(wp(before=120, after=40) + wr("*POSSIBILIDADE DE PARCELAMENTO*", bold=True) + "</w:p>")
            xml.append(wp(before=0, after=40) + wr("Parcelamento Simples Nacional") + "</w:p>")
            if not b.get("omitir_total"):
                xml.append(wp(before=0, after=40) + wr("Valor total parcelado: ") + wr(b.get("valor_total",""), bold=True) + wr(".") + "</w:p>")
            vp = b.get("valor_primeira",""); vd = b.get("valor_demais",""); np_ = b.get("num_parcelas",0)
            if vp and vp != vd:
                xml.append(wp_bullet(normal="Entrada de ", bold_txt=f"{vp} + {np_-1}x de {vd}."))
            else:
                xml.append(wp_bullet(normal="", bold_txt=f"{np_}x de {vd}."))
        elif t == "parc_simp":
            xml.append(wp(before=120, after=40) + wr("*POSSIBILIDADE DE PARCELAMENTO*", bold=True) + "</w:p>")
            xml.append(wp(before=0, after=40) + wr("Parcelamento Simplificado (e-CAC)") + "</w:p>")
            xml.append(wp(before=0, after=40) + wr(f'Principal: {b.get("principal","")} | Multa: {b.get("multa","")} | Juros: {b.get("juros","")}') + "</w:p>")
            xml.append(wp(before=0, after=40) + wr("Valor total parcelado: ") + wr(b.get("total",""), bold=True) + wr(".") + "</w:p>")
            xml.append(wp_bullet(normal="", bold_txt=f'{b.get("num_parcelas","")}x de {b.get("valor_parcela","")}.'  ))
        elif t == "parc_siefpar":
            atraso = b.get("parcelas_em_atraso", 0)
            xml.append(wp(before=100, after=40) + wr("Parcelamento em andamento — nº ", bold=True) + wr(b.get("numero",""), bold=True) + "</w:p>")
            xml.append(wp_bullet(normal="", bold_txt=f'{atraso} PARCELA{"S" if atraso > 1 else ""} EM ATRASO — {b.get("valor_em_atraso","")}'))
    return "".join(xml)


# ══════════════════════════════════════════════════════════
# 4. MONTAGEM DOS TEXTOS
# ══════════════════════════════════════════════════════════

def montar_federal(federal, parc_sn, parc_simp, parc_dau, hoje, parc_sispar, reg_valores, decl_omissas,
                    pgdau_prest=None, mei_emissao=None, analise_previa=None, pgmei=None):
    blocos = []
    # Alerta de indeferimento do Simples Nacional (Relatório de Análise Prévia) —
    # vem primeiro por ser a informação mais urgente do memorando quando presente.
    if analise_previa and analise_previa.get("indeferida"):
        data = f" a partir de {analise_previa['data_vigencia']}" if analise_previa.get("data_vigencia") else ""
        blocos.append({"type":"alerta",
            "text": f"⚠ ATENÇÃO: a opção pelo Simples Nacional será INDEFERIDA{data} devido aos "
                    f"débitos abaixo, caso não sejam regularizados."})
        blocos.append({"type":"separator"})
    if federal.get("certidao_numero"):
        val = (f", válida até {federal['certidao_validade']}" if federal.get("certidao_validade") else "")
        rotulo = federal.get("certidao_tipo") or "Certidão Positiva com Efeitos de Negativa"
        blocos.append({"type":"normal_bold","text":f"{rotulo} gerada","bold":f"{val}."})

    # Declarações omissas
    if decl_omissas and decl_omissas.get("encontrado"):
        blocos.append({"type":"label","text":"Declarações em atraso:"})
        for decl in decl_omissas["declaracoes"]:
            blocos.append({"type":"bullet","normal":f"{decl['tipo']} — ","bold":decl["descricao"]})
        # Multa
        multa_un = decl_omissas["declaracoes"][0]["multa_unitaria"] if decl_omissas["declaracoes"] else 200.0
        multa_fmt = fmt(multa_un)
        blocos.append({"type":"normal",
            "text": f"Valor da multa: {multa_fmt} por declaração "
                    f"(Redução de 50% desse valor se pago dentro do prazo)."})

    # Parcelamento com exigibilidade suspensa (PARCSN/PARCMEI)
    if federal.get("parcelamento_ativo"):
        blocos.append({"type":"date",
            "text": f"Parcelamento com Exigibilidade Suspensa — {federal['parcelamento_ativo']} (débitos suspensos, não exigíveis)"})

    # Parcelas em atraso de parcelamentos já ativos — MEI (eCAC) e Dívida Ativa (PGFN)
    tem_mei_atraso = mei_emissao and mei_emissao.get("encontrado")
    tem_dau_atraso = pgdau_prest and pgdau_prest.get("encontrado") and pgdau_prest.get("valor_parcela_atual", 0) > 0
    tem_pgmei_atraso = pgmei and pgmei.get("encontrado") and pgmei.get("meses_em_atraso")
    if tem_mei_atraso or tem_dau_atraso or tem_pgmei_atraso:
        blocos.append({"type":"label","text":"Débitos:"})
        if tem_mei_atraso:
            blocos.append({"type":"bullet",
                "normal": f"PARCELAMENTO MEI ATRASO – {mei_emissao['quantidade']} parcelas de {fmt(mei_emissao['valor_parcela'])} totalizando ",
                "bold": fmt(mei_emissao['total']) + "."})
        if tem_dau_atraso:
            blocos.append({"type":"bullet",
                "normal": f"PARCELAMENTO DÍVIDA ATIVA – parcela do mês {pgdau_prest['mes_parcela_atual']} de ",
                "bold": fmt(pgdau_prest['valor_parcela_atual']) +
                        (f" (Restando {pgdau_prest['parcelas_restantes']} parcelas a vencer)"
                         if pgdau_prest.get("parcelas_restantes", 0) > 0 else "") + "."})
        if tem_pgmei_atraso:
            meses = pgmei["meses_em_atraso"]
            qtd = len(meses)
            faixa = (f"{meses[0]['mes']}/{meses[0]['ano']} a {meses[-1]['mes']}/{meses[-1]['ano']}"
                     if qtd > 1 else f"{meses[0]['mes']}/{meses[0]['ano']}")
            blocos.append({"type":"bullet",
                "normal": f"DAS MEI EM ATRASO – {faixa} ({qtd} {plural(qtd, 'guia', 'guias')}) totalizando ",
                "bold": fmt(pgmei['total_atraso']) + "."})

    # Anos do MEI ainda não liberados no PGMEI (dependem da entrega da DASN
    # do ano exibido) — não dá pra saber o valor exato (falta multa/juros,
    # que dependem da data de pagamento), então mostra um mínimo estimado
    # (guias pendentes × valor padrão do DAS, sem multa/juros/encargos).
    if pgmei and pgmei.get("anos_em_aberto"):
        for info in pgmei["anos_em_aberto"]:
            valor_min = f" Valor mínimo estimado: {fmt(info['valor_minimo'])}." if info.get("valor_minimo") else ""
            blocos.append({"type":"date",
                "text": f"* DAS MEI {info['ano']}: em aberto — liberado após a entrega da DASN "
                        f"{pgmei['ano_calendario']} ({info['quantidade_guias']} "
                        f"{plural(info['quantidade_guias'], 'guia pendente', 'guias pendentes')})."
                        f"{valor_min}"})

    if federal.get("debitos"):
        por_receita = {}
        for d in federal["debitos"]:
            por_receita.setdefault(d["receita"],[]).append(d)
        # Quando o Parcelamento do Simples Nacional (PGDAS) cobre a única receita em
        # aberto, o valor mostrado é o do parcelamento — o que vai ser efetivado —
        # em vez do saldo devedor bruto do SIEF (confirmado com memorando real).
        usar_valor_parcelamento_sn = (
            parc_sn.get("encontrado") and not parc_sn.get("indisponivel") and len(por_receita) == 1
        )
        for receita, items in por_receita.items():
            periodos = [i["periodo"] for i in items]
            faixa = f"{periodos[0]} a {periodos[-1]}" if len(periodos)>1 else periodos[0]
            total = parc_sn["valor_total"] if usar_valor_parcelamento_sn else sum(i["saldo_devedor"] for i in items)
            blocos.append({"type":"italic","text":"Referente à guias do imposto do " + receita.rstrip(".") + "."})
            blocos.append({"type":"bullet","normal":f"Referente ao mês {faixa} – ","bold":fmt(total)})

        if parc_sn.get("encontrado") and not parc_sn.get("indisponivel"):
            blocos.append({"type":"parc_sn","valor_total":fmt(parc_sn["valor_total"]),"num_parcelas":parc_sn["num_parcelas"],"valor_primeira":fmt(parc_sn["valor_primeira"]),"valor_demais":fmt(parc_sn["valor_demais"]),"omitir_total":usar_valor_parcelamento_sn})
        elif parc_sn.get("indisponivel"):
            blocos.append({"type":"date","text":"* Parcelamento via Simples Nacional não disponível para este caso (valores abaixo do mínimo exigido)."})

        if parc_simp.get("encontrado") and not parc_simp.get("indisponivel"):
            blocos.append({"type":"parc_simp","principal":fmt(parc_simp["principal"]),"multa":fmt(parc_simp["multa"]),"juros":fmt(parc_simp["juros"]),"total":fmt(parc_simp["total"]),"num_parcelas":parc_simp["num_parcelas"],"valor_parcela":fmt(parc_simp["valor_parcela"])})
        elif parc_simp.get("indisponivel"):
            blocos.append({"type":"date","text":"* Parcelamento Simplificado indisponível (dívida abaixo do mínimo necessário)."})

        if federal.get("parcelamento_siefpar"):
            sp = federal["parcelamento_siefpar"]
            blocos.append({"type":"parc_siefpar","numero":sp["numero"],"parcelas_em_atraso":sp["parcelas_em_atraso"],"valor_em_atraso":fmt(sp["valor_em_atraso"])})

        blocos.append({"type":"separator"})

    # Dívida Ativa: usar Regularize_valores se disponível, senão SIDA do SITUAÇÃO FISCAL
    sida = federal.get("inscricoes_sida", [])
    if reg_valores and reg_valores.get("encontrado") and reg_valores.get("inscricoes"):
        # Usar valores do Regularize (mais preciso)
        rv_insc = reg_valores["inscricoes"]
        total_rv = reg_valores["total"]
        # Detectar categoria (primeira inscrição do SIDA ou genérico)
        cat_sida = sida[0]["receita"] if sida else "PGFN"
        blocos.append({"type":"label",
            "text": f"Dívida Ativa — {cat_sida} ({len(rv_insc)} {plural(len(rv_insc), 'inscrição', 'inscrições')})"})
        for insc in rv_insc:
            prot = " — ⚠ protestada" if insc.get("protesto") else ""
            blocos.append({"type":"bullet",
                "normal": f"{insc['inscricao']} — ",
                "bold":   fmt(insc["valor"]) + prot})
        blocos.append({"type":"normal_bold",
            "text": "Total: ",
            "bold": fmt(total_rv)})
    elif sida and not parc_dau.get("encontrado"):
        por_data = {}
        for insc in sida:
            por_data.setdefault(insc["data"], []).append(insc)
        for data, inscricoes in por_data.items():
            receitas = ", ".join(i["receita"] for i in inscricoes)
            total_insc = len(inscricoes)
            blocos.append({"type":"label",
                "text": f"Dívida Ativa — {total_insc} {plural(total_insc, 'inscrição', 'inscrições')} na PGFN"})
            blocos.append({"type":"normal", "text": f"Receitas: {receitas}"})
            blocos.append({"type":"normal_bold",
                "text": f"Inscritas em {data} — situação: ",
                "bold": inscricoes[0].get("situacao","ATIVA EM COBRANÇA")})

    if federal.get("exigibilidade_suspensa"):
        por_receita_es = {}
        for d in federal["exigibilidade_suspensa"]:
            por_receita_es.setdefault(d["receita"],[]).append(d)
        for receita, items in por_receita_es.items():
            total = sum(i["saldo_devedor"] for i in items)
            periodos = [i["periodo"] for i in items]
            faixa = f"{periodos[0]} a {periodos[-1]}" if len(periodos)>1 else periodos[0]
            blocos.append({"type":"italic","text":"Débito com Exigibilidade Suspensa — " + receita.rstrip(".") + "."})
            blocos.append({"type":"bullet","normal":f"Período {faixa} — ","bold":fmt(total)+" (a analisar/vencer)"})
        blocos.append({"type":"separator"})

    if parc_dau.get("encontrado"):
        # Formatar lista de inscrições se for relatório consolidado
        inscricoes_fmt = []
        for insc in parc_dau.get("inscricoes", []):
            inscricoes_fmt.append({
                "inscricao": insc["inscricao"],
                "situacao":  insc["situacao"],
                "valor":     insc["valor"]
            })
        # Inscrições com situação "NEGOCIADA" já estão em parcelamento ativo (ex.: SISPAR) —
        # não são dívida em aberto à espera de negociação, e o texto não pode dar a entender isso.
        if inscricoes_fmt:
            qtd_negociadas = sum(1 for i in inscricoes_fmt if "NEGOCIADA" in i["situacao"].upper())
            situacao_negociada = qtd_negociadas > 0
            todas_negociadas = qtd_negociadas == len(inscricoes_fmt)
        else:
            situacao_negociada = "NEGOCIADA" in parc_dau.get("situacao","").upper()
            todas_negociadas = situacao_negociada
        blocos.append({
            "type":      "divida",
            "receita":   parc_dau.get("receita",""),
            "inscricao": parc_dau.get("inscricao",""),
            "data":      parc_dau.get("data_inscricao",""),
            "situacao":  parc_dau.get("situacao",""),
            "situacao_negociada": situacao_negociada,
            "todas_negociadas":   todas_negociadas,
            "principal": fmt(parc_dau.get("principal",0)),
            "multa":     fmt(parc_dau.get("multa",0)),
            "juros":     fmt(parc_dau.get("juros",0)),
            "encargo":   fmt(parc_dau.get("encargo",0)),
            "total":     fmt(parc_dau.get("total",0)),
            "protesto":  parc_dau.get("protesto",False),
            "inscricoes": inscricoes_fmt
        })

    # SISPAR — opções de parcelamento PGFN
    if parc_sispar and parc_sispar.get("encontrado") and parc_sispar.get("opcoes"):
        opcoes_com = [o for o in parc_sispar["opcoes"] if o.get("disponivel") and o.get("total_a_pagar", 0) > 0]
        opcoes_sem = [o for o in parc_sispar["opcoes"] if not o.get("disponivel")]

        if opcoes_com or opcoes_sem:
            blocos.append({"type":"label","text":"*POSSIBILIDADE DE PARCELAMENTO — PGFN (SISPAR)*"})

        for op in opcoes_com:
            cat    = op.get("categoria","")
            np_    = op.get("num_prestacoes", 0)
            vp_    = fmt(op.get("valor_prestacao", 0))
            total_ = fmt(op.get("total_a_pagar", 0))
            nd_    = op.get("num_dividas", 0)
            blocos.append({"type":"italic","text":f"{cat} — {nd_} {plural(nd_, 'inscrição negociável', 'inscrições negociáveis')}"})
            if np_ == 1:
                blocos.append({"type":"bullet","normal":"Parcelamento Convencional: ","bold":f"1x de {vp_}"})
            else:
                blocos.append({"type":"bullet","normal":"Parcelamento Convencional: ",
                               "bold":f"{np_}x de {vp_} — Total: {total_}"})

        for op in opcoes_sem:
            cat = op.get("categoria","")
            blocos.append({"type":"date","text":f"* {cat}: sem opção de negociação disponível no momento."})

    # Uma única data "valores atualizados", sempre por último — não uma por
    # sub-seção (o texto tinha data repetida no meio, sem sentido pra quem lê).
    if blocos:
        blocos.append({"type":"date","text":f"(valores atualizados no dia {hoje})"})
    else:
        blocos.append({"type":"normal","text":"Sem débitos. CND gerada."})
    return blocos_para_xml(blocos)


def montar_estadual(estadual, site_contrib=None, icms=None, ipva=None, hoje=None):
    blocos = []
    tem_icms = bool(icms and icms.get("encontrado") and icms.get("total", 0) > 0)
    # IPVA não entra nessa checagem de propósito: no memorando validado que serviu
    # de referência, a seção de IPVA aparece sozinha, sem a linha "Débitos
    # encontrados" — diferente do ICMS, que já foi confirmado COM essa linha.
    tem_ipva = bool(ipva and ipva.get("encontrado") and ipva.get("total", 0) > 0)
    if not estadual.get("tem_debitos") and not tem_icms:
        blocos.append({"type":"normal","text":"Sem débitos. CND gerada."})
        for c in estadual.get("certidaos",[]):
            orgao = c.get("orgao",""); num = c.get("numero",""); val = c.get("validade","")
            prefix = "CRDA" if "PGE" in orgao else "CND"
            val_txt = f" — válida até {val}" if val else ""
            blocos.append({"type":"normal_bold","text":f"{prefix} nº ","bold":f"{num} ({orgao}){val_txt}: Sem débitos."})
    else:
        blocos.append({"type":"normal","text":"Débitos encontrados. Verificar certidões."})

    # Débitos de IPVA — Secretaria da Fazenda e Planejamento do Estado de São Paulo
    if tem_ipva:
        blocos.append({"type":"label","text":"IPVA:"})
        for v in ipva["veiculos"]:
            blocos.append({"type":"normal",
                "text": f"Placa {v['placa']} — RENAVAM {v['renavam']} — Exercício {v['exercicio']}"})
        blocos.append({"type":"normal_bold","text":"Valor total: ","bold": fmt(ipva["total"]) + "."})

    # Simulação de Parcelamento ICMS — Site do Contribuinte (SEFAZ-SP)
    if tem_icms:
        blocos.append({"type":"label","text":"*POSSIBILIDADE DE PARCELAMENTO — Dívida Ativa: ICMS DEVIDO*"})
        qtd_icms = icms['quantidade_dividas']
        blocos.append({"type":"italic",
            "text": f"ICMS Devido — {qtd_icms} {plural(qtd_icms, 'inscrição negociável', 'inscrições negociáveis')}"})
        if icms.get("num_parcelas"):
            blocos.append({"type":"bullet","normal":"Parcelamento Convencional: ",
                "bold": f"{icms['num_parcelas']}x de {fmt(icms['valor_parcela'])} — Total: {fmt(icms['total'])}"})
        else:
            blocos.append({"type":"bullet","normal":"Total: ","bold": fmt(icms['total'])})

    # Site do Contribuinte — pendências GIA/EFD
    if site_contrib and site_contrib.get("encontrado") and site_contrib.get("pendencias"):
        sc = site_contrib
        total = sc["total_competencias"]
        p_ini = sc["primeira_omissao"]
        p_fim = sc["ultima_omissao"]
        faixa = f"{p_ini} a {p_fim}" if p_ini != p_fim else p_ini
        data_primeira = sc.get("primeira_data","")

        blocos.append({"type":"separator"})
        blocos.append({"type":"label",
            "text":"Pendências Estaduais — SEFAZ-SP (Site do Contribuinte)"})
        blocos.append({"type":"italic",
            "text":"GIA/EFD Omissa — declarações não entregues:"})
        blocos.append({"type":"bullet",
            "normal": f"Períodos {faixa} — ",
            "bold":   f"{total} competência{'s' if total > 1 else ''} omissa{'s' if total > 1 else ''}"})
        if data_primeira:
            blocos.append({"type":"normal",
                "text": f"Primeira omissão registrada desde: {data_primeira}"})
        blocos.append({"type":"date",
            "text": "(pendências de entrega de declaração — não representam débito financeiro)"})

    # Uma única data "valores atualizados", sempre por último — mesmo padrão
    # usado no federal/municipal (mostrar só quando há valor que possa mudar).
    if estadual.get("tem_debitos") or tem_icms or tem_ipva:
        blocos.append({"type":"date","text":f"(valores atualizados no dia {hoje})"})

    return blocos_para_xml(blocos)


# Tabela de preços do escritório — lista completa de serviços, para consulta
# manual e como fonte única de valores para TABELA_HONORARIOS abaixo. Nem
# todo item tem extrator automático ainda. Cada valor é (preço, unidade) —
# preço None quando o próprio preço é uma fórmula (ex.: "30% do recuperado")
# em vez de um valor fixo.
TABELA_PRECOS_ESCRITORIO = {
    "XEROX/IMPRESSÃO":                                    (1.50,    "UN"),
    "RETIFICAÇÃO DAS MEI":                                (10.00,   "UN"),
    "IMPOSTO DE RENDA":                                   (120.00,  "A PARTIR DE"),
    "RETIFICAÇÃO DE GUIA DO IMPOSTO DE RENDA":            (20.00,   "A PARTIR DE"),
    "PEDIDO DE COMPRA":                                   (70.00,   "A PARTIR DE"),
    "GUIA DE INSS AVULSA":                                (15.00,   "UN"),
    "ORÇAMENTO":                                          (100.00,  "A PARTIR DE"),
    "RETIFICAÇÃO DE IMPOSTO DE RENDA":                    (100.00,  "A PARTIR DE"),
    "CERTIFICADO DIGITAL A1":                             (280.00,  "UN"),
    "CERTIFICADO DIGITAL A3":                             (550.00,  "UN"),
    "CONSULTA E REPARCELAMENTO DE DÉBITOS":               (60.00,   "A PARTIR DE"),
    "RETIFICAÇÃO/ALTERAÇÃO DE CONTRATO":                  (50.00,   "A PARTIR DE"),
    "RETIFICAÇÃO DE GUIA DE PARCELAMENTO":                (20.00,   "UN"),
    "RETIFICAÇÃO DE RESCISÃO":                            (50.00,   "UN"),
    "RETIFICAÇÃO DAS SIMPLES":                            (20.00,   "UN"),
    "RETIFICAÇÃO FGTS":                                   (20.00,   "UN"),
    "RETIFICAÇÃO INSS":                                   (20.00,   "UN"),
    "RETIFICAÇÃO IRRF":                                   (20.00,   "UN"),
    "RETIFICAÇÃO GRRF":                                   (25.00,   "UN"),
    "RETIFICAÇÃO ICMS":                                   (20.00,   "UN"),
    "EMISSÃO NOTA SERVIÇOS":                              (30.00,   "UN"),
    "EMISSÃO NOTA PRODUTO":                               (35.00,   "A PARTIR DE"),
    "PACOTE 1 MENSAL (10 NOTAS FISCAIS) - EXCEDENTE R$ 30,00": (200.00, "A PARTIR DE"),
    "PACOTE 2 MENSAL (15 NOTAS FISCAIS) - EXCEDENTE R$ 30,00": (270.00, "A PARTIR DE"),
    "PACOTE 3 MENSAL (20 NOTAS FISCAIS) - EXCEDENTE R$ 30,00": (320.00, "A PARTIR DE"),
    "EMISSÃO CTE":                                        (50.00,   "A PARTIR DE"),
    "CONTRATO DE ALUGUEL":                                (130.00,  "A PARTIR DE"),
    "CONTRATO DE COMPRA/VENDA":                           (160.00,  "A PARTIR DE"),
    "KIT PLACAS OBRIGATORIAS":                            (60.00,   "UN"),
    "LIVRO DE OCORRENCIA DP":                             (40.00,   "UN"),
    "LIVRO DE HORAS DP/FISCAL":                           (40.00,   "UN"),
    "LICITAÇÃO (EDITAL/DOCS)":                            (600.00,  "A PARTIR DE"),
    "AVCB/CLCB BOMBEIRO":                                 (750.00,  "A PARTIR DE"),
    "LEITORA DE CERTIFICADO DIGITAL":                     (180.00,  "UN"),
    "ABERTURA DE CNPJ ME":                                (1380.00, "A PARTIR DE"),
    "BAIXA DE CNPJ ME":                                   (1680.00, "A PARTIR DE"),
    "ALTERAÇÃO DE CNPJ ME":                                (1380.00, "A PARTIR DE"),
    "TRANSFORMAÇÃO DE ME PARA LTDA":                      (1680.00, "A PARTIR DE"),
    "ABERTURA DE CNPJ MEI":                               (250.00,  "A PARTIR DE"),
    "BAIXA DE CNPJ MEI":                                  (250.00,  "A PARTIR DE"),
    "ALTERAÇÃO DE CNPJ MEI":                               (250.00,  "A PARTIR DE"),
    "REGISTRO DE MARCA E PATENTE":                        (2800.00, "A PARTIR DE"),
    "REGISTROS DE PJ EM CONSELHOS DE CLASSE":             (550.00,  "A PARTIR DE"),
    "REGISTRO DE PJ NO CPOM":                             (550.00,  "A PARTIR DE"),
    "DECLARAÇÃO ANUAL MEI":                               (150.00,  "A PARTIR DE"),
    "DECLARAÇÃO DEFIS ANUAL SIMPLES NACIONAL":            (250.00,  "A PARTIR DE"),
    "DECLARAÇÃO DE FATURAMENTO AVULSA":                   (100.00,  "A PARTIR DE"),
    "TREINAMENTO DE EMISSÃO DE NF":                       (220.00,  "A PARTIR DE"),
    "HORA DA CONSULTA C/CONTADOR":                        (280.00,  "HORA"),
    "PGDAS MENSAL":                                       (65.00,   "A PARTIR DE"),
    "EFD-CONTRIB.":                                       (120.00,  "A PARTIR DE"),
    "DCTF":                                               (40.00,   "A PARTIR DE"),
    "GFIP ZERADA":                                        (40.00,   "A PARTIR DE"),
    "ITR":                                                (300.00,  "A PARTIR DE"),
    "CAR":                                                (400.00,  "A PARTIR DE"),
    "DAP":                                                (380.00,  "A PARTIR DE"),
    "TALÃO DE NOTAS DO PRODUTOR":                         (200.00,  "UN"),
    "CCIR / INCRA":                                       (380.00,  "A PARTIR DE"),
    "CERTIDÃO DE MATRICULA DE IMOVEIS RURAL":             (180.00,  "A PARTIR DE"),
    "EMISSÃO DE GTA RURAL":                               (45.00,   "A PARTIR DE"),
    "INTERMEDIAR ACORDO TRABALHISTA":                     (600.00,  "A PARTIR DE"),
    "ABERTURA DE PRODUTOR RURAL":                         (750.00,  "A PARTIR DE"),
    "ALTERAÇÃO CNPJ RURAL":                               (600.00,  "A PARTIR DE"),
    "DECLARAÇÃO DE VACINAÇÃO DE GADO":                    (150.00,  "A PARTIR DE"),
    "RECUPERAÇÃO TRIBUTARIA":                             (None,    "30% DO VALOR RECUPERADO"),
    "SUSPENÇÃO DE CNPJ":                                  (1080.00, "A PARTIR DE"),
    "GCAP (GANHO DE CAPITAL) CARRO":                      (200.00,  "A PARTIR DE"),
    "GCAP (GANHO DE CAPITAL) CASA":                       (380.00,  "A PARTIR DE"),
    "GCAP (GANHO DE CAPITAL) PROPRIEDADE RURAL":          (750.00,  "A PARTIR DE"),
    "PARCELAMENTO DAS MEI":                               (150.00,  "A PARTIR DE"),
    "DCTF WEB SEM MOVIMENTO":                             (180.00,  "A PARTIR DE"),
    "DCTF COM MOVIMENTO":                                 (200.00,  "A PARTIR DE"),
    "CNO DE OBRAS":                                       (None,    "30% DO VALOR ECONOMIZADO"),
    "ENTRADA AUXILIO MATERNDADE MEI":                     (250.00,  "A PARTIR DE"),
    "ABERTURA DE HOLDING ZERADA":                         (5000.00, "A PARTIR DE"),
    "ABERTURA DE HOLDING COM INTEGRALIZAÇÃO":             (5000.00, "30% DA ECONOMIA"),
    "ABERTURA DE ASSOCIAÇÃO SEM FINS LUCRATIVOS":         (3000.00, "A PARTIR DE"),
    "CONSULTORIA EM REFORMA TRIBUTÁRIA":                  (1621.00, "A PARTIR DE"),
    "COMISSÃO DE SISTEMA NF-E/SAT":                       (None,    "VERIFICAR SISTEMA"),
}

# Subconjunto de TABELA_PRECOS_ESCRITORIO já usado no cálculo automático de
# honorários (calcular_honorarios) — puxa o valor de lá para não haver dois
# lugares com o mesmo preço podendo divergir.
TABELA_HONORARIOS = {
    "DECLARACAO_ANUAL_MEI":        TABELA_PRECOS_ESCRITORIO["DECLARAÇÃO ANUAL MEI"][0],              # por ano
    "DCTF_WEB_SEM_MOVIMENTO":      TABELA_PRECOS_ESCRITORIO["DCTF WEB SEM MOVIMENTO"][0],            # sempre 1 (ao fazer uma, demais somem)
    "DCTF":                         TABELA_PRECOS_ESCRITORIO["DCTF"][0],                              # por ano
    "EFD_CONTRIB":                 TABELA_PRECOS_ESCRITORIO["EFD-CONTRIB."][0],                      # por ano
    "PARCELAMENTO_MEI":            TABELA_PRECOS_ESCRITORIO["PARCELAMENTO DAS MEI"][0],              # único
    "CONSULTA_REPARCELAMENTO":      TABELA_PRECOS_ESCRITORIO["CONSULTA E REPARCELAMENTO DE DÉBITOS"][0],  # único
    "RETIFICACAO_DAS_MEI":          TABELA_PRECOS_ESCRITORIO["RETIFICAÇÃO DAS MEI"][0],              # por guia (mês em atraso no PGMEI, ou mês pendente de ano ainda não liberado)
}

def calcular_honorarios(federal, decl_omissas, municipal, parc_sn, parc_simp, parc_dau, parc_sispar, reg_valores, pgmei=None):
    """Calcula automaticamente os honorários com base nas pendências do cliente."""
    itens = []  # [{"descricao": str, "qtd": int, "valor_unit": float, "total": float}]

    # ── Declarações omissas ─────────────────────────────────────
    if decl_omissas and decl_omissas.get("declaracoes"):
        for decl in decl_omissas["declaracoes"]:
            tipo = decl["tipo"]

            if tipo == "DASN SIMEI":
                qtd = len(decl["periodos"])
                unit = TABELA_HONORARIOS["DECLARACAO_ANUAL_MEI"]
                itens.append({
                    "descricao": f"Declaração Anual MEI ({qtd}x)",
                    "qtd": qtd, "valor_unit": unit, "total": qtd * unit
                })

            elif tipo == "DCTFWeb":
                # Sempre 1 — ao fazer a primeira, as demais somem
                unit = TABELA_HONORARIOS["DCTF_WEB_SEM_MOVIMENTO"]
                itens.append({
                    "descricao": "DCTFWeb sem movimento (1x)",
                    "qtd": 1, "valor_unit": unit, "total": unit
                })

            elif tipo == "DCTF":
                qtd = len(decl["periodos"])
                unit = TABELA_HONORARIOS["DCTF"]
                itens.append({
                    "descricao": f"DCTF ({qtd}x)",
                    "qtd": qtd, "valor_unit": unit, "total": qtd * unit
                })

            elif tipo == "EFD-CONTRIBUIÇÕES":
                qtd = len(decl["periodos"])
                unit = TABELA_HONORARIOS["EFD_CONTRIB"]
                itens.append({
                    "descricao": f"EFD-Contribuições ({qtd}x)",
                    "qtd": qtd, "valor_unit": unit, "total": qtd * unit
                })

    # ── Dívida Ativa MEI (SIDA com SIMPLES NACIONAL MEI) ────────
    sida = federal.get("inscricoes_sida", []) if federal else []
    tem_sida_mei = any("MEI" in i.get("receita","").upper() for i in sida)
    tem_sida_geral = bool(sida) and not tem_sida_mei

    if tem_sida_mei or (reg_valores and reg_valores.get("encontrado")):
        unit = TABELA_HONORARIOS["PARCELAMENTO_MEI"]
        itens.append({
            "descricao": "Parcelamento das MEI",
            "qtd": 1, "valor_unit": unit, "total": unit
        })
    elif tem_sida_geral:
        unit = TABELA_HONORARIOS["CONSULTA_REPARCELAMENTO"]
        itens.append({
            "descricao": "Consulta e Reparcelamento de Débitos",
            "qtd": 1, "valor_unit": unit, "total": unit
        })

    # ── Parcelamento Simplificado (e-CAC) ────────────────────────
    if parc_simp and parc_simp.get("encontrado") and not parc_simp.get("indisponivel"):
        unit = TABELA_HONORARIOS["CONSULTA_REPARCELAMENTO"]
        itens.append({
            "descricao": "Consulta e Reparcelamento (Simplificado e-CAC)",
            "qtd": 1, "valor_unit": unit, "total": unit
        })

    # ── SISPAR / Regularize ──────────────────────────────────────
    if parc_sispar and parc_sispar.get("opcoes"):
        opcoes_com_valor = [o for o in parc_sispar["opcoes"] if o.get("total_a_pagar", 0) > 0]
        if opcoes_com_valor:
            unit = TABELA_HONORARIOS["CONSULTA_REPARCELAMENTO"]
            itens.append({
                "descricao": "Consulta e Reparcelamento PGFN (SISPAR)",
                "qtd": 1, "valor_unit": unit, "total": unit
            })

    # ── Débitos municipais ───────────────────────────────────────
    if municipal and municipal.get("tem_debitos"):
        unit = TABELA_HONORARIOS["CONSULTA_REPARCELAMENTO"]
        itens.append({
            "descricao": "Consulta e Reparcelamento Municipal",
            "qtd": 1, "valor_unit": unit, "total": unit
        })

    # ── PGMEI — DAS mensal do MEI em atraso + meses de anos ainda não
    # liberados (pendentes da entrega da DASN) — cobrado por guia, cada uma
    # precisando de retificação/emissão avulsa.
    if pgmei and pgmei.get("encontrado"):
        qtd_atraso = len(pgmei.get("meses_em_atraso", []))
        qtd_em_aberto = sum(a["quantidade_guias"] for a in pgmei.get("anos_em_aberto", []))
        qtd = qtd_atraso + qtd_em_aberto
        if qtd:
            unit = TABELA_HONORARIOS["RETIFICACAO_DAS_MEI"]
            itens.append({
                "descricao": f"Retificação DAS MEI ({qtd}x)",
                "qtd": qtd, "valor_unit": unit, "total": qtd * unit
            })

    # Deduplicar CONSULTA_REPARCELAMENTO (cobrar só uma vez mesmo com múltiplas fontes)
    visto_reparc = False
    itens_dedup = []
    for item in itens:
        if "Reparcelamento" in item["descricao"] or "reparcelamento" in item["descricao"]:
            if visto_reparc:
                continue
            visto_reparc = True
        itens_dedup.append(item)

    total = sum(i["total"] for i in itens_dedup)
    return {"itens": itens_dedup, "total": total}


def montar_resumo(federal, decl_omissas, municipal, parc_siefpar, honorarios, hoje,
                  parc_sn=None, parc_simp=None, parc_dau=None, parc_sispar=None, reg_valores=None, pgmei=None):
    """Gera o bloco de resumo final do memorando."""
    blocos = []
    blocos.append({"type":"label","text":"RESUMO"})

    # Calcular honorários automaticamente se não passado manualmente
    if honorarios == 0 and any([parc_sn, parc_simp, parc_dau, parc_sispar, reg_valores,
                                 decl_omissas and decl_omissas.get("encontrado"),
                                 municipal and municipal.get("tem_debitos"),
                                 pgmei and pgmei.get("encontrado")]):
        calc = calcular_honorarios(
            federal, decl_omissas, municipal,
            parc_sn, parc_simp, parc_dau, parc_sispar, reg_valores, pgmei
        )
        honorarios_calc = calc["total"]
        itens_honorarios = calc["itens"]
    else:
        honorarios_calc = honorarios
        itens_honorarios = [{"descricao": "Honorários contábeis", "total": honorarios}] if honorarios > 0 else []

    # Federal: multas + parcelamento
    total_fed = 0.0
    partes_fed = []

    multas_total = decl_omissas.get("total_multas", 0.0) if decl_omissas else 0.0
    if multas_total > 0:
        partes_fed.append(f"{fmt(multas_total)} (com possibilidade de redução)")
        total_fed += multas_total

    # SIEFPAR
    if parc_siefpar:
        saldo_atraso = parc_siefpar.get("valor_em_atraso", 0.0)
        if saldo_atraso > 0:
            partes_fed.append(f"{fmt(saldo_atraso)} parcelas")
            total_fed += saldo_atraso

    # PGMEI — DAS MEI em atraso (valor real) + mínimo estimado do(s) ano(s)
    # ainda em aberto (mesmo critério do corpo do texto: soma tudo que é
    # dinheiro devido, não só o que já tem valor 100% fechado).
    if pgmei and pgmei.get("encontrado"):
        if pgmei.get("total_atraso", 0) > 0:
            partes_fed.append(fmt(pgmei["total_atraso"]))
            total_fed += pgmei["total_atraso"]
        valor_min_aberto = sum(a.get("valor_minimo", 0.0) for a in pgmei.get("anos_em_aberto", []))
        if valor_min_aberto > 0:
            partes_fed.append(fmt(valor_min_aberto))
            total_fed += valor_min_aberto

    if partes_fed:
        blocos.append({"type":"bullet",
            "normal": "Âmbito Federal: ",
            "bold":   " + ".join(partes_fed)})

    # Municipal — mesma prioridade usada no corpo do texto: quando há um
    # levantamento específico do que entra no parcelamento, o total do
    # resumo tem que bater com o total mostrado no âmbito, não com a
    # certidão sozinha.
    total_mun = 0.0
    if municipal:
        total_mun = municipal.get("total_parcelamento", 0.0) or municipal.get("total", 0.0)
    if total_mun > 0:
        blocos.append({"type":"bullet",
            "normal": "Âmbito Municipal: ",
            "bold":   fmt(total_mun)})

    # Honorários — detalhar por item se calculado automaticamente
    if itens_honorarios:
        if len(itens_honorarios) == 1:
            blocos.append({"type":"bullet",
                "normal": "Honorários contábeis para regularização: ",
                "bold":   fmt(honorarios_calc)})
        else:
            blocos.append({"type":"label","text":"Honorários contábeis para regularização:"})
            for item in itens_honorarios:
                blocos.append({"type":"bullet",
                    "normal": f"{item['descricao']}: ",
                    "bold":   fmt(item["total"])})
            blocos.append({"type":"normal_bold",
                "text": "Subtotal honorários: ",
                "bold": fmt(honorarios_calc)})

    # Total geral
    total_geral = total_fed + total_mun + honorarios_calc
    obs = " (restando apenas parcelas a vencer)" if parc_siefpar and parc_siefpar.get("valor_em_atraso", 0) > 0 else ""
    blocos.append({"type":"normal_bold",
        "text": "VALOR TOTAL PARA REGULARIZAÇÃO: ",
        "bold": fmt(total_geral) + obs})

    return blocos_para_xml(blocos)


def montar_municipal(municipal, hoje):
    blocos = []
    if not municipal.get("tem_debitos"):
        blocos.append({"type":"normal","text":"Sem débitos. CND gerada."})
    else:
        # Certidão Positiva Municipal (Fiorilli)
        if municipal.get("certidao_numero"):
            blocos.append({"type":"normal_bold",
                "text": "Certidão Positiva Municipal gerada",
                "bold": "."})

        # Quando existe um levantamento específico dos débitos que entram no
        # parcelamento ("Listagem de Débito"/Formato 4), ele tem prioridade
        # sobre a certidão — é o valor que efetivamente vai ser parcelado,
        # não só o retrato da dívida no momento da emissão da certidão.
        usar_parcelamento = municipal.get("total_parcelamento", 0) > 0
        debitos_fonte = municipal["debitos_parcelamento"] if usar_parcelamento else municipal.get("debitos", [])
        total_fonte = municipal["total_parcelamento"] if usar_parcelamento else municipal.get("total", 0)

        # Agrupar débitos por receita
        agrupados = {}
        for d in debitos_fonte:
            chave = f"{d['receita']} {d['ano']} ({d['status']})" if d.get("status") else f"{d['receita']} {d['ano']}"
            agrupados.setdefault(chave, {"parcelas":[], "a_pagar":0.0})
            if d.get("parcela"):
                agrupados[chave]["parcelas"].append(d["parcela"])
            agrupados[chave]["a_pagar"] += d.get("a_pagar", 0.0)

        for chave, dados in agrupados.items():
            parcelas = dados["parcelas"]
            if not parcelas:
                desc = ""
            else:
                desc = f"parcela {parcelas[0]}" if len(parcelas)==1 else f"parcelas {', '.join(parcelas)}"
            prefixo = f"{chave} — {desc} — A pagar: " if desc else f"{chave} — A pagar: "
            if dados["a_pagar"] > 0:
                blocos.append({"type":"bullet","normal":prefixo, "bold":fmt(dados["a_pagar"])})
            else:
                blocos.append({"type":"bullet","normal":"","bold":f"{chave} — {desc}" if desc else chave})

        if total_fonte > 0:
            blocos.append({"type":"normal_bold","text":"Total a pagar: ","bold":fmt(total_fonte)})

        # Opções de parcelamento (Simulação de Débitos — Fiorilli): mostra a opção
        # com mais parcelas (menor valor mensal), igual ao critério usado nos demais
        # parcelamentos do memorando.
        opcoes = municipal.get("parcelamento_opcoes")
        if opcoes:
            melhor = max(opcoes, key=lambda o: o["num_parcelas"])
            qtd = len(municipal.get("debitos", []))
            blocos.append({"type":"label","text":"*POSSIBILIDADE DE PARCELAMENTO — Débitos da prefeitura*"})
            blocos.append({"type":"italic","text":f"Débitos — {qtd} {plural(qtd, 'inscrição negociável', 'inscrições negociáveis')}"})
            blocos.append({"type":"bullet","normal":"Parcelamento Convencional: ",
                "bold": f"{melhor['num_parcelas']}x de {fmt(melhor['demais'])} — Total: {fmt(municipal.get('total', melhor['demais']*melhor['num_parcelas']))}"})
        else:
            blocos.append({"type":"date","text":"*Não há possibilidade de parcelamento*."})

        # Data por último, no lugar de onde ficava o placeholder antigo do modelo.
        blocos.append({"type":"date","text":f"(valores atualizados no dia {hoje})"})
    return blocos_para_xml(blocos)


# ══════════════════════════════════════════════════════════
# 5. PREENCHIMENTO DO .DOCX
# ══════════════════════════════════════════════════════════

def texto_para_xml_runs(texto, rpr_base=""):
    linhas = str(texto).split("\n")
    runs = []
    for i, linha in enumerate(linhas):
        if i > 0: runs.append(f'<w:r>{rpr_base}<w:br/></w:r>')
        if linha:
            esc = escape_xml(linha)
            sp  = ' xml:space="preserve"' if linha != linha.strip() else ""
            runs.append(f'<w:r>{rpr_base}<w:t{sp}>{esc}</w:t></w:r>')
    return "".join(runs)


def substituir_placeholder_xml(xml, placeholder, valor):
    tag = "{{" + placeholder + "}}"
    if tag not in xml: return xml
    vs = str(valor)
    if vs.strip().startswith("<w:p"):
        def rp(m):
            return vs if tag in m.group(0) else m.group(0)
        return re.sub(r'<w:p[ >].*?</w:p>', rp, xml, flags=re.DOTALL)
    if "\n" not in vs:
        return xml.replace(tag, escape_xml(vs))
    pattern = re.compile(r'<w:r>((?:<w:rPr>.*?</w:rPr>)?)<w:t[^>]*>' + re.escape(tag) + r'</w:t></w:r>', re.DOTALL)
    resultado = pattern.sub(lambda m: texto_para_xml_runs(vs, m.group(1)), xml)
    if tag in resultado:
        resultado = resultado.replace(f"<w:t>{tag}</w:t>", texto_para_xml_runs(vs))
    return resultado


def preencher_docx(template_path, output_path, dados):
    shutil.copy2(template_path, output_path)
    tmp = output_path + ".tmp"
    alvos = {"word/document.xml","word/header1.xml","word/header2.xml","word/footer1.xml"}
    with zipfile.ZipFile(output_path, 'r') as zin:
        with zipfile.ZipFile(tmp, 'w', zipfile.ZIP_DEFLATED) as zout:
            for item in zin.infolist():
                data = zin.read(item.filename)
                if item.filename in alvos:
                    try:
                        xml = data.decode('utf-8')
                        for chave, valor in dados.items():
                            xml = substituir_placeholder_xml(xml, chave, valor)
                        data = xml.encode('utf-8')
                    except Exception as e:
                        print(f"AVISO: {item.filename}: {e}", file=sys.stderr)
                zout.writestr(item, data)
    os.replace(tmp, output_path)
    print(f"Documento gerado: {output_path}")


# ══════════════════════════════════════════════════════════
# 6. MAIN
# ══════════════════════════════════════════════════════════

def processar_pasta(pasta_pdfs, honorarios=0.0):
    """Classifica e extrai os dados de todos os PDFs de uma pasta (modo
    subpastas federal/estadual/municipal, ou pasta única com classificação
    automática) e devolve tudo pronto para preencher o modelo — sem gravar
    nenhum arquivo. Usado pelo main() (CLI) e pela interface web, para que
    os dois caminhos rodem exatamente a mesma lógica de extração.
    """
    TIPOS_FEDERAL = {"federal","parcelamento_sn","parc_sn_indisponivel","parcelamento_simplificado","parcelamento_dau","parcelamento_sispar","regularize_valores","parcelamento_pgdau_prestacoes","parcelamento_mei_emissao","analise_previa_sn","mei_pgmei"}
    textos = {}
    classificacao = []  # [(nome_exibido, tipo), ...] — para revisão antes de gerar

    tem_sub = any(os.path.isdir(os.path.join(pasta_pdfs, s)) for s in ["federal","estadual","municipal"])

    if tem_sub:
        mapa = {"federal":TIPOS_FEDERAL, "estadual":{"estadual_pge","estadual_sefaz","estadual_site_contribuinte","estadual_icms_parcelamento","estadual_ipva"}, "municipal":{"municipal"}}
        for sub, tipos_validos in mapa.items():
            caminho_sub = os.path.join(pasta_pdfs, sub)
            if not os.path.isdir(caminho_sub): continue
            for arq in sorted(os.listdir(caminho_sub)):
                if not arq.lower().endswith(".pdf"): continue
                caminho = os.path.join(caminho_sub, arq)
                texto   = extrair_texto_pdf(caminho)
                tipo    = classificar_pdf(texto)
                if tipo not in tipos_validos:
                    tipo = list(tipos_validos)[0]
                textos.setdefault(tipo, []).append(texto)
                classificacao.append((f"{sub}/{arq}", tipo))
    else:
        for arq in sorted(os.listdir(pasta_pdfs)):
            if not arq.lower().endswith(".pdf"): continue
            texto = extrair_texto_pdf(os.path.join(pasta_pdfs, arq))
            tipo  = classificar_pdf(texto)
            textos.setdefault(tipo, []).append(texto)
            classificacao.append((arq, tipo))

    def get(tipo): return "\n\n".join(textos.get(tipo,[])) or ""

    nome, cnpj = extrair_dados_empresa({k: get(k) for k in textos})
    federal     = extrair_federal(get("federal"))
    decl_omissas   = extrair_declaracoes_omissas(get("federal"))
    parc_sn     = extrair_parcelamento_sn(get("parcelamento_sn") or get("parc_sn_indisponivel"))
    parc_simp   = extrair_parcelamento_simplificado(get("parcelamento_simplificado"))
    parc_dau    = extrair_parcelamento_dau(get("parcelamento_dau"))
    parc_sispar     = extrair_parcelamento_sispar(get("parcelamento_sispar"))
    reg_valores     = extrair_regularize_valores(get("regularize_valores"))
    pgdau_prest = extrair_pgdau_prestacoes(get("parcelamento_pgdau_prestacoes"))
    analise_previa = extrair_analise_previa_sn(get("analise_previa_sn"))
    mei_emissao = extrair_mei_emissao_parcela(get("parcelamento_mei_emissao"))
    pgmei       = extrair_pgmei(get("mei_pgmei"))
    estadual    = extrair_estadual(get("estadual_pge"), get("estadual_sefaz"))
    site_contrib = extrair_site_contribuinte(get("estadual_site_contribuinte"))
    icms_parc   = extrair_icms_parcelamento(get("estadual_icms_parcelamento"))
    ipva        = extrair_ipva(get("estadual_ipva"))
    municipal   = extrair_municipal(get("municipal"))
    hoje        = date.today().strftime("%d.%m.%Y")

    siefpar = federal.get("parcelamento_siefpar")

    dados_template = {
        "NOME_EMPRESA":             nome or "EMPRESA NÃO IDENTIFICADA",
        "CNPJ":                     cnpj or "00.000.000/0000-00",
        "DATA_CONSULTA":            date.today().strftime("%d/%m/%Y"),
        "FEDERAL_TEXTO":            montar_federal(federal, parc_sn, parc_simp, parc_dau, hoje, parc_sispar, reg_valores, decl_omissas, pgdau_prest, mei_emissao, analise_previa, pgmei),
        "ESTADUAL_TEXTO":           montar_estadual(estadual, site_contrib, icms_parc, ipva, hoje),
        "MUNICIPAL_TEXTO":          montar_municipal(municipal, hoje),
        "RESUMO":                   montar_resumo(federal, decl_omissas, municipal, siefpar, honorarios, hoje, parc_sn, parc_simp, parc_dau, parc_sispar, reg_valores, pgmei),
        "DATA_ATUALIZACAO_VALORES": date.today().strftime("%d/%m/%Y"),
    }

    dados_json = {"empresa":{"nome":nome,"cnpj":cnpj},"federal":federal,"parcelamento_sn":parc_sn,"parcelamento_simplificado":parc_simp,"parcelamento_dau":parc_dau,"parcelamento_sispar":parc_sispar,"regularize_valores":reg_valores,"pgdau_prestacoes":pgdau_prest,"mei_emissao":mei_emissao,"pgmei":pgmei,"analise_previa_sn":analise_previa,"declaracoes_omissas":decl_omissas,"estadual":estadual,"site_contribuinte":site_contrib,"icms_parcelamento":icms_parc,"ipva":ipva,"municipal":municipal}

    resumo = {
        "empresa": nome, "cnpj": cnpj,
        "federal_status": federal["status"],
        "parc_sn": ("SIM (indisponível)" if parc_sn.get("indisponivel") else "SIM") if parc_sn.get("encontrado") else "não",
        "parc_simplificado": "SIM" if parc_simp.get("encontrado") else "não",
        "parc_dau": "SIM" if parc_dau.get("encontrado") else "não",
        "parc_sispar": [o["categoria"] for o in parc_sispar.get("opcoes",[]) if o.get("disponivel")],
        "siefpar_atraso": federal.get("parcelamento_siefpar"),
        "estadual_status": estadual["status"],
        "site_contribuinte_omissoes": site_contrib.get("total_competencias") if site_contrib.get("encontrado") else None,
        "municipal_status": municipal["status"],
        "analise_previa_indeferida": analise_previa.get("indeferida", False),
    }

    return {
        "nome": nome, "cnpj": cnpj,
        "classificacao": classificacao,
        "dados_template": dados_template,
        "dados_json": dados_json,
        "resumo": resumo,
    }


def main():
    if len(sys.argv) < 4:
        print("Uso: python processar_fiscal.py <pasta_pdfs> <template.docx> <output.docx>")
        sys.exit(1)

    pasta_pdfs, template, output = sys.argv[1], sys.argv[2], sys.argv[3]

    # Honorários: parâmetro opcional na linha de comando
    honorarios = 0.0
    if len(sys.argv) >= 5:
        try:
            honorarios = float(sys.argv[4].replace(",", "."))
        except:
            pass

    modo = "Modo subpastas detectado." if any(
        os.path.isdir(os.path.join(pasta_pdfs, s)) for s in ["federal","estadual","municipal"]
    ) else "Modo pasta única."
    print(modo)

    r = processar_pasta(pasta_pdfs, honorarios)
    for arq, tipo in r["classificacao"]:
        print(f"  {arq} → {tipo}")

    s = r["resumo"]
    print("\n=== RESUMO ===")
    print(f"Empresa        : {s['empresa']}")
    print(f"CNPJ           : {s['cnpj']}")
    print(f"Federal        : {s['federal_status']}")
    print(f"Parc. SN       : {s['parc_sn']}")
    print(f"Parc. Simplif. : {s['parc_simplificado']}")
    print(f"Parc. DAU      : {s['parc_dau']}")
    print(f"Parc. SISPAR   : {'SIM - ' + ', '.join(s['parc_sispar']) if s['parc_sispar'] else 'não'}")
    print(f"SIEFPAR atraso : {s['siefpar_atraso'] or 'não'}")
    print(f"Estadual       : {s['estadual_status']}")
    print(f"Site Contrib.  : {'SIM - ' + str(s['site_contribuinte_omissoes']) + ' omissões' if s['site_contribuinte_omissoes'] else 'não'}")
    print(f"Municipal      : {s['municipal_status']}")
    print()

    os.makedirs(os.path.dirname(os.path.abspath(output)), exist_ok=True)
    preencher_docx(template, output, r["dados_template"])

    json_path = output.replace(".docx","_dados.json")
    with open(json_path, "w", encoding="utf-8") as f:
        json.dump(r["dados_json"], f, ensure_ascii=False, indent=2)
    print(f"JSON exportado: {json_path}")

if __name__ == "__main__":
    main()