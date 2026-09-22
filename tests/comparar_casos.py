#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Compara o memorando gerado pelo script com o memorando validado, caso a caso.

Estrutura esperada (a pasta casos/ NÃO é versionada):
  casos/<GRUPO>/<CASO>/*.pdf          PDFs de origem
  casos/<GRUPO>/<CASO>/<memorando>.docx   memorando validado (nome começa com "memorando")

Uso:
  python tests/comparar_casos.py [--template modelo.docx] [--casos casos] [--filtro TEXTO] [-v]

Código de saída: 0 se todos os casos baterem, 1 se houver divergência.
"""
import argparse, difflib, os, re, shutil, subprocess, sys, tempfile, zipfile
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
SCRIPT = RAIZ / "scripts" / "processar_fiscal.py"
TEMPLATES_PADRAO = [RAIZ / "templates" / "modelo_memorando.docx", RAIZ / "MODELO_MEMORANDO_n8n.docx"]

DATA = r"\d{2}[./]\d{2}[./]\d{4}"


def paragrafos_docx(caminho):
    """Texto de cada parágrafo do corpo do .docx, na ordem."""
    xml = zipfile.ZipFile(caminho).read("word/document.xml").decode("utf-8")
    saida = []
    for p in re.findall(r"<w:p[ >].*?</w:p>", xml, flags=re.DOTALL):
        p = re.sub(r"<w:br\s*/>", "\n", p)
        p = re.sub(r"<w:tab\s*/>", " ", p)
        texto = "".join(re.findall(r"<w:t[^>]*>(.*?)</w:t>", p, flags=re.DOTALL))
        texto = (texto.replace("&amp;", "&").replace("&lt;", "<").replace("&gt;", ">")
                      .replace("&quot;", '"').replace("&apos;", "'"))
        saida.extend(texto.split("\n"))
    return saida


def normalizar(linhas):
    """Remove diferenças que não importam (bullets, espaços, datas de emissão)."""
    saida = []
    for l in linhas:
        l = l.replace(" ", " ")
        l = re.sub(r"^[\s•·▪●\-–]+(?=\S)", "", l) if re.match(r"^\s*[•·▪●]", l) else l
        l = re.sub(r"\s+", " ", l).strip()
        if not l:
            continue
        l = re.sub(rf"(valores atualizados no dia ){DATA}", r"\1<DATA>", l, flags=re.I)
        l = re.sub(rf"(Consulta realizada no dia ){DATA}", r"\1<DATA>", l, flags=re.I)
        if re.fullmatch(DATA, l):
            l = "<DATA>"
        saida.append(l)
    return saida


def achar_template(arg):
    if arg:
        return Path(arg)
    for t in TEMPLATES_PADRAO:
        if t.exists():
            return t
    sys.exit("Nenhum modelo encontrado. Use --template.")


def descobrir_casos(pasta, filtro):
    for grupo in sorted(p for p in Path(pasta).iterdir() if p.is_dir()):
        for caso in sorted(p for p in grupo.iterdir() if p.is_dir()):
            if filtro and filtro.lower() not in caso.name.lower():
                continue
            docx = [f for f in caso.glob("*.docx") if f.name.lower().startswith("memorando") and not f.name.startswith("~$")]
            pdfs = [f for f in caso.glob("*.pdf") if not f.name.lower().startswith("memorando")]
            if len(docx) != 1 or not pdfs:
                print(f"[PULADO] {grupo.name}/{caso.name}: {len(docx)} memorando(s) .docx, {len(pdfs)} PDF(s) de origem")
                continue
            yield grupo.name, caso.name, docx[0], pdfs


def rodar_caso(pdfs, template, tmp):
    entrada = Path(tmp) / "pdfs"
    entrada.mkdir()
    for f in pdfs:
        shutil.copy2(f, entrada / f.name)
    saida = Path(tmp) / "gerado.docx"
    r = subprocess.run([sys.executable, str(SCRIPT), str(entrada), str(template), str(saida)],
                       capture_output=True, text=True, encoding="utf-8", errors="replace")
    return saida, r


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("--casos", default=str(RAIZ / "casos"))
    ap.add_argument("--template")
    ap.add_argument("--filtro")
    ap.add_argument("-v", "--verbose", action="store_true", help="mostra o texto do script (classificação dos PDFs)")
    a = ap.parse_args()

    template = achar_template(a.template)
    total = ok = 0
    for grupo, nome, docx_ok, pdfs in descobrir_casos(a.casos, a.filtro):
        total += 1
        with tempfile.TemporaryDirectory() as tmp:
            gerado, r = rodar_caso(pdfs, template, tmp)
            print(f"\n{'=' * 78}\n{grupo} / {nome}")
            if a.verbose or r.returncode != 0:
                print(r.stdout.strip()); print(r.stderr.strip())
            if r.returncode != 0 or not gerado.exists():
                print(f"  ✗ script falhou (código {r.returncode})")
                continue
            esperado = normalizar(paragrafos_docx(docx_ok))
            obtido = normalizar(paragrafos_docx(gerado))
        if esperado == obtido:
            ok += 1
            print(f"  ✓ idêntico ({len(esperado)} linhas)")
            continue
        dif = [d for d in difflib.unified_diff(esperado, obtido, "validado", "gerado", n=1, lineterm="")]
        print(f"  ✗ {sum(1 for d in dif if d[:1] in '+-' and d[:3] not in ('+++', '---'))} linha(s) divergente(s)"
              f"  (validado={len(esperado)} linhas, gerado={len(obtido)})")
        for d in dif:
            print("   ", d)
    print(f"\n{'=' * 78}\nResultado: {ok}/{total} casos idênticos")
    sys.exit(0 if ok == total and total > 0 else 1)


if __name__ == "__main__":
    main()
