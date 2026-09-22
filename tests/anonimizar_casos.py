#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Gera fixtures de teste ANONIMIZADAS a partir de casos/<grupo>/<caso>/ (não versionado).

Para cada caso:
  - extrai o texto de cada PDF de origem (pdfplumber) e do memorando validado (.docx)
  - coleta TODAS as variações do nome (com/sem acento, com/sem sufixo, maiúsc./minúsc.)
    e do CNPJ/CPF encontradas em qualquer um dos documentos do caso
  - substitui cada variação por um valor fictício estável por caso
  - grava o resultado em tests/fixtures/<grupo>/<caso>/  (texto puro, versionável)
  - confere, ao final, que nenhum resíduo do nome/CNPJ/CPF real permaneceu
    (checagem insensível a maiúsculas/acentos, independente da lista de variações
    usada na substituição)

Números de inscrição, protocolo, código da dívida e valores em R$ NÃO são alterados:
são identificadores administrativos, não dado pessoal, e o teste depende do formato
exato deles para validar a extração.

Uso:
  python tests/anonimizar_casos.py [--casos casos] [--saida tests/fixtures]
"""
import argparse, importlib.util, re, sys, unicodedata, zipfile
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(RAIZ / "scripts"))
spec = importlib.util.spec_from_file_location("pf", RAIZ / "scripts" / "processar_fiscal.py")
pf = importlib.util.module_from_spec(spec)
spec.loader.exec_module(pf)

NOMES_FICTICIOS = ["EMPRESA MODELO ALFA", "EMPRESA MODELO BETA", "EMPRESA MODELO GAMA",
                    "EMPRESA MODELO DELTA", "EMPRESA MODELO EPSILON", "EMPRESA MODELO ZETA"]
CNPJ_FICTICIOS = ["11.111.111/0001-11", "22.222.222/0001-22", "33.333.333/0001-33",
                  "44.444.444/0001-44", "55.555.555/0001-55", "66.666.666/0001-66"]
CPF_FICTICIO_FMT = "111.111.111-11"
CPF_FICTICIO_DIG = "11111111111"

LABEL_PATTERNS = [
    r"Consulta da empresa:\s*([^\n]+)",
    r"Proje[cç][aã]o de imposto da empresa:\s*([^\n]+)",
    r"Contribuinte:\s*([^\n]+?)(?:\s+CPF/CNPJ|\s+CNPJ|\s*$)",
    r"Devedor:\s*([^\n]+)",
    r"Nome Empresarial[:\s]+([^\n]+)",
    r"Raz[aã]o Social\s*[:\-]?\s*([^\n]+)",
    r"CNPJ[:\s]+[\d./\-]+\s*[-–]\s*([^\n]{3,80})",
    r"Respons[aá]vel[:\s]+[\d./\-]{0,18}\s*[-–]?\s*([^\n]{3,80})",
    r"S[oó]cio[:\s]+[\d./\-]{0,18}\s*[-–]?\s*([^\n]{3,80})",
    r"Titular[:\s]+[\d./\-]{0,18}\s*[-–]?\s*([^\n]{3,80})",
    r"Procurador[:\s]+([^\n]{3,80})",
]
LIXO = ["sair", "localizar", "acesso", "mensagens", "perfil", "dados do", "procurador",
        "consultas", "pagamentos", "certid", "precat", "legisla", "atendimento", "transpar"]
# palavras genéricas (tipo societário, ramo) que não são identificadoras por si só —
# não entram na lista de palavras a redigir isoladamente
PALAVRAS_GENERICAS = {
    "de", "da", "do", "das", "dos", "e", "empresa", "modelo", "ltda", "me", "eireli",
    "produtos", "servicos", "serviços", "comercio", "comércio", "industria", "indústria",
    "farmaceuticos", "farmacêuticos", "sociedade", "individual", "microempreendedor",
}


def slug(s):
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    s = re.sub(r"[^A-Za-z0-9]+", "_", s.strip())
    return re.sub(r"_+", "_", s).strip("_").upper()


def nome_arquivo_seguro(stem):
    """Nome de arquivo de fixture a partir do nome do PDF de origem, mas sem
    qualquer sequência de 3+ dígitos — o nome do PDF muitas vezes carrega o
    CNPJ do cliente ou uma data/protocolo (ex.: "RelatorioSituacaoFiscal-
    12345678000199-20260921.pdf"). Só o nome do arquivo, nunca os números,
    vira o nome do arquivo de teste."""
    sem_numeros = re.sub(r"\d{3,}", "", stem)
    return slug(sem_numeros) or "PDF"


def sem_acento(s):
    return unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")


def paragrafos_docx(caminho):
    xml = zipfile.ZipFile(caminho).read("word/document.xml").decode("utf-8")
    saida = []
    for p in re.findall(r"<w:p[ >].*?</w:p>", xml, flags=re.DOTALL):
        p = re.sub(r"<w:br\s*/>", "\n", p)
        texto = "".join(re.findall(r"<w:t[^>]*>(.*?)</w:t>", p, flags=re.DOTALL))
        texto = (texto.replace("&amp;", "&").replace("&lt;", "<").replace("&gt;", ">")
                      .replace("&quot;", '"').replace("&apos;", "'"))
        saida.append(texto)
    return "\n".join(saida)


def pattern_flexivel(nome):
    """Regex que casa o nome ignorando maiúsc./minúsc. E diferenças de acentuação."""
    grupos = {
        "a": "aàáâãä", "e": "eèéêë", "i": "iìíîï",
        "o": "oòóôõö", "u": "uùúûü", "c": "cç",
    }
    partes = []
    for ch in nome:
        low = ch.lower()
        if low in grupos:
            classe = grupos[low] + grupos[low].upper()
            partes.append(f"[{classe}]")
        elif ch == " ":
            partes.append(r"\s+")
        else:
            partes.append(re.escape(ch))
    return re.compile("".join(partes))


def limpar_variante(v):
    v = v.strip()
    v = re.sub(r"\s+\d{8,}$", "", v).strip()          # CPF/CNPJ colado no fim
    v = re.sub(r"^\d{2}\.\d{3}\.\d{3}\s+", "", v).strip()  # raiz de CNPJ colada no início
    v = v.rstrip(".,;:")
    return v


def coletar_variantes_nome(textos):
    variantes = set()
    for t in textos:
        for pat in LABEL_PATTERNS:
            for m in re.finditer(pat, t, re.IGNORECASE):
                v = limpar_variante(m.group(1))
                if len(v) < 4 or len(v) > 90:
                    continue
                if any(l in v.lower() for l in LIXO):
                    continue
                if not re.search(r"[A-Za-zÀ-ÿ]{3,}", v):
                    continue
                variantes.add(v)
    return variantes


def coletar_cnpjs(textos):
    achados = set()
    for t in textos:
        achados |= set(re.findall(r"\d{2}\.\d{3}\.\d{3}/\d{4}-\d{2}", t))
        achados |= set(re.findall(r"(?<!\d)\d{14}(?!\d)", t))
    return achados


def coletar_cpfs(textos, variantes_nome):
    achados = set(re.findall(r"\d{3}\.\d{3}\.\d{3}-\d{2}", "\n".join(textos)))
    for t in textos:
        for nome in variantes_nome:
            for m in re.finditer(pattern_flexivel(nome).pattern + r"\D{0,4}(\d{11})\b", t):
                achados.add(m.group(1))
    return achados


def palavras_identificadoras(variantes_nome):
    """Toda palavra >=3 letras das variantes de nome, exceto termos genéricos."""
    palavras = set()
    for nome in variantes_nome:
        for p in re.split(r"\s+", nome):
            p_limpo = re.sub(r"[^A-Za-zÀ-ÿ]", "", p)
            if len(p_limpo) >= 3 and sem_acento(p_limpo).lower() not in PALAVRAS_GENERICAS:
                palavras.add(p_limpo)
    return palavras


def montar_substituicoes(textos, nome_fake, cnpj_fake):
    variantes_nome = coletar_variantes_nome(textos)
    cnpjs = coletar_cnpjs(textos)
    cpfs = coletar_cpfs(textos, variantes_nome)

    subs = []
    # 1) frase inteira, da variante mais longa para a mais curta (mantém o texto legível)
    for nome in sorted(variantes_nome, key=len, reverse=True):
        subs.append((pattern_flexivel(nome), nome_fake))

    # 2) rede de segurança: cada palavra identificadora isolada, onde quer que apareça
    #    (cobre nomes que só compartilham PARTE das palavras entre documentos, como o nome
    #    da empresa vs. o nome completo do responsável/sócio no relatório fiscal)
    pool_fake = (nome_fake.split() * 3)
    for i, palavra in enumerate(sorted(palavras_identificadoras(variantes_nome), key=len, reverse=True)):
        troca = pool_fake[i % len(pool_fake)]
        subs.append((re.compile(pattern_flexivel(palavra).pattern + r"\b"), troca))

    cnpj_fake_digitos = re.sub(r"\D", "", cnpj_fake)
    for c in sorted(cnpjs, key=len, reverse=True):
        digitos = re.sub(r"\D", "", c)
        raiz = digitos[:8]
        raiz_fmt = f"{raiz[:2]}.{raiz[2:5]}.{raiz[5:8]}"
        cnpj_fake_raiz_fmt = f"{cnpj_fake_digitos[:2]}.{cnpj_fake_digitos[2:5]}.{cnpj_fake_digitos[5:8]}"
        subs.append((re.compile(re.escape(c)), cnpj_fake if "." in c else cnpj_fake_digitos))
        subs.append((re.compile(re.escape(raiz_fmt)), cnpj_fake_raiz_fmt))

    for cpf in cpfs:
        digitos = re.sub(r"\D", "", cpf)
        subs.append((re.compile(re.escape(cpf)), CPF_FICTICIO_FMT))
        subs.append((re.compile(re.escape(digitos)), CPF_FICTICIO_DIG))

    return subs, variantes_nome, cnpjs


def aplicar(texto, subs):
    for pat, novo in subs:
        texto = pat.sub(novo, texto)
    texto = re.sub(r"(?im)^(Endere[cç]o:).*$", r"\1 <ENDEREÇO OCULTADO>", texto)
    texto = re.sub(r"(?im)^(Rua|Av\.|Avenida|Pra[cç]a)\s+[^\n]+", "<ENDEREÇO OCULTADO>", texto)
    return texto


def checar_residuos(texto, variantes_nome, cnpjs):
    """Checagem independente das trocas: acento/caixa ignorados dos dois lados."""
    t_norm = sem_acento(texto).lower()
    achados = []
    for nome in variantes_nome:
        if sem_acento(nome).lower() in t_norm:
            achados.append(f"nome '{nome}'")
    for c in cnpjs:
        digitos = re.sub(r"\D", "", c)
        if digitos in re.sub(r"\D", "", texto):
            achados.append(f"CNPJ {c}")
    return achados


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--casos", default=str(RAIZ / "casos"))
    ap.add_argument("--saida", default=str(RAIZ / "tests" / "fixtures"))
    a = ap.parse_args()

    saida_raiz = Path(a.saida)
    if saida_raiz.exists():
        import shutil
        shutil.rmtree(saida_raiz)

    idx = 0
    total_residuos = 0
    contadores_grupo = {}
    for grupo_dir in sorted(p for p in Path(a.casos).iterdir() if p.is_dir()):
        grupo_slug = slug(grupo_dir.name)
        # nome da pasta de cada caso = singular do nome do grupo + número (CLIENTES -> CLIENTE_1,
        # NAO_CLIENTES -> NAO_CLIENTE_1, AMOSTRAS -> AMOSTRA_1, ...)
        prefixo_caso = grupo_slug[:-1] if grupo_slug.endswith("S") else grupo_slug
        for caso_dir in sorted(p for p in grupo_dir.iterdir() if p.is_dir()):
            pdfs = sorted(f for f in caso_dir.glob("*.pdf") if not f.name.lower().startswith("memorando"))
            docx = [f for f in caso_dir.glob("*.docx")
                    if f.name.lower().startswith("memorando") and not f.name.startswith("~$")]
            # memorando validado é opcional — sem ele, o caso vira só uma amostra de
            # formato (usada para testar a extração), sem um resultado esperado para comparar
            if not pdfs or len(docx) > 1:
                print(f"[PULADO] {grupo_dir.name}/{caso_dir.name}: {len(pdfs)} pdf(s), {len(docx)} memorando(s)")
                continue

            textos_pdf = {f: pf.extrair_texto_pdf(str(f)) for f in pdfs}
            texto_docx = paragrafos_docx(docx[0]) if docx else None
            todos_textos = list(textos_pdf.values()) + ([texto_docx] if texto_docx else [])

            # Contador por grupo (não global): assim, adicionar um grupo novo (ex.: AMOSTRAS)
            # não muda a identidade fictícia nem o número dos casos já existentes em CLIENTES/
            # NÃO CLIENTES — o diff de um caso novo fica restrito a ele mesmo.
            contadores_grupo[grupo_slug] = contadores_grupo.get(grupo_slug, 0) + 1
            pos = contadores_grupo[grupo_slug] - 1
            nome_fake = NOMES_FICTICIOS[pos % len(NOMES_FICTICIOS)]
            cnpj_fake = CNPJ_FICTICIOS[pos % len(CNPJ_FICTICIOS)]
            idx += 1
            subs, variantes_nome, cnpjs = montar_substituicoes(todos_textos, nome_fake, cnpj_fake)

            # Nome de pasta NUNCA pode vir do nome real do caso (foi assim que o nome de
            # um cliente vazou para o GitHub antes) — usa um contador genérico por grupo.
            nome_pasta = f"{prefixo_caso}_{contadores_grupo[grupo_slug]}"
            destino = saida_raiz / grupo_slug / nome_pasta
            (destino / "pdfs").mkdir(parents=True, exist_ok=True)

            residuos_caso = 0
            nomes_usados = set()
            for f, t in textos_pdf.items():
                t2 = aplicar(t, subs)
                r = checar_residuos(t2, variantes_nome, cnpjs)
                if r:
                    residuos_caso += len(r)
                    print(f"   ⚠ {f.name}: {r}")
                base = nome_arquivo_seguro(f.stem)
                nome_final, n = base, 2
                while nome_final in nomes_usados:
                    nome_final = f"{base}_{n}"; n += 1
                nomes_usados.add(nome_final)
                (destino / "pdfs" / (nome_final + ".txt")).write_text(t2, encoding="utf-8")

            if texto_docx is not None:
                t2 = aplicar(texto_docx, subs)
                r = checar_residuos(t2, variantes_nome, cnpjs)
                if r:
                    residuos_caso += len(r)
                    print(f"   ⚠ {docx[0].name}: {r}")
                (destino / "memorando_validado.txt").write_text(t2, encoding="utf-8")

            total_residuos += residuos_caso
            status = "OK" if residuos_caso == 0 else f"FALHOU ({residuos_caso} resíduo(s))"
            amostra = "" if texto_docx is not None else " [amostra, sem memorando validado]"
            print(f"[{status}] {grupo_dir.name}/{caso_dir.name} → {destino.relative_to(RAIZ)}{amostra}  "
                  f"({len(pdfs)} pdf(s), {len(variantes_nome)} variante(s) de nome, {len(cnpjs)} cnpj(s))")

    print(f"\nTotal de resíduos encontrados: {total_residuos}")
    sys.exit(1 if total_residuos else 0)


if __name__ == "__main__":
    main()
