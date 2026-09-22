#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Teste de regressão por FATOS (não por texto exato) contra as fixtures
anonimizadas em tests/fixtures/ — não depende de casos/ nem de dado real,
por isso pode rodar em qualquer máquina, sem PDF nem memorando de cliente.

Compara valores que já foram conferidos manualmente, um a um, contra os
PDFs reais originais (ver histórico de commits para a evidência de cada
um). Uma quebra aqui quer dizer que uma mudança no script alterou um
resultado que já sabemos que está certo — não necessariamente que o
resultado NOVO está errado, mas merece checar antes de seguir.

Uso:
  python tests/test_fatos.py [-v]

Código de saída: 0 se tudo bateu, 1 se algo quebrou.
"""
import argparse, glob, importlib.util, sys
from pathlib import Path

RAIZ = Path(__file__).resolve().parent.parent
sys.path.insert(0, str(RAIZ / "scripts"))
spec = importlib.util.spec_from_file_location("pf", RAIZ / "scripts" / "processar_fiscal.py")
pf = importlib.util.module_from_spec(spec)
spec.loader.exec_module(pf)


def carregar_caso(caminho):
    """Classifica cada PDF do caso e agrupa o texto por tipo, como main() faz."""
    textos = {}
    for f in sorted(Path(caminho).glob("pdfs/*.txt")):
        t = f.read_text(encoding="utf-8")
        tipo = pf.classificar_pdf(t)
        textos.setdefault(tipo, []).append(t)
    return {k: "\n\n".join(v) for k, v in textos.items()}


def extrair_fatos(t):
    """Roda todos os extratores relevantes e devolve os números que o
    memorando final usa — é o que este teste compara, não o texto."""
    get = lambda k: t.get(k, "")
    decl = pf.extrair_declaracoes_omissas(get("federal"))
    dau = pf.extrair_parcelamento_dau(get("parcelamento_dau"))
    sispar = pf.extrair_parcelamento_sispar(get("parcelamento_sispar"))
    icms = pf.extrair_icms_parcelamento(get("estadual_icms_parcelamento"))
    municipal = pf.extrair_municipal(get("municipal"))
    pgdau = pf.extrair_pgdau_prestacoes(get("parcelamento_pgdau_prestacoes"))
    mei = pf.extrair_mei_emissao_parcela(get("parcelamento_mei_emissao"))
    estadual = pf.extrair_estadual(get("estadual_pge"), get("estadual_sefaz"))
    federal = pf.extrair_federal(get("federal"))
    parc_sn = pf.extrair_parcelamento_sn(get("parcelamento_sn"))
    analise = pf.extrair_analise_previa_sn(get("analise_previa_sn"))
    return {
        "multas_total": decl["total_multas"],
        "multas_qtd": len(decl["declaracoes"]),
        "dau_total": dau["total"],
        "dau_inscricoes": len(dau["inscricoes"]),
        "sispar_opcoes": len(sispar["opcoes"]),
        "icms_total": icms["total"],
        "icms_quantidade_dividas": icms["quantidade_dividas"],
        "municipal_total": municipal["total"],
        "municipal_debitos": len(municipal["debitos"]),
        "municipal_total_parcelamento": municipal["total_parcelamento"],
        "pgdau_valor_atual": pgdau["valor_parcela_atual"],
        "pgdau_restantes": pgdau["parcelas_restantes"],
        "mei_total": mei["total"],
        "mei_quantidade": mei["quantidade"],
        "estadual_certidoes": len(estadual["certidaos"]),
        "federal_certidao_tipo": federal["certidao_tipo"],
        "federal_debitos": len(federal["debitos"]),
        "parc_sn_total": parc_sn["valor_total"],
        "parc_sn_num_parcelas": parc_sn["num_parcelas"],
        "analise_previa_indeferida": analise["indeferida"],
    }


# Valores conferidos manualmente contra os PDFs reais (anonimizados aqui).
# Só entra aqui o que já foi validado contra um memorando real ou contra o
# próprio PDF de origem — ver o histórico de commits para a evidência.
CASOS_ESPERADOS = {
    "CLIENTES/CLIENTE_1": {
        "sispar_opcoes": 6,
        "icms_total": 971.15,
        "icms_quantidade_dividas": 2,
        "municipal_total": 938.43,
        "municipal_debitos": 5,
    },
    "CLIENTES/CLIENTE_2": {
        "sispar_opcoes": 1,
        "estadual_certidoes": 1,
    },
    "NAO_CLIENTES/NAO_CLIENTE_1": {
        "dau_total": 3305.51,
        "dau_inscricoes": 2,
    },
    "NAO_CLIENTES/NAO_CLIENTE_2": {
        "multas_total": 1000.0,
        "multas_qtd": 4,
        "dau_total": 7606.69,
        "dau_inscricoes": 2,
        "pgdau_valor_atual": 134.93,
        "pgdau_restantes": 26,
        "mei_total": 1166.22,
        "mei_quantidade": 19,
    },
    "AMOSTRAS/AMOSTRA_1": {
        "federal_certidao_tipo": "Certidão Negativa",
        "federal_debitos": 7,
        "parc_sn_total": 37986.19,
        "parc_sn_num_parcelas": 60,
        "municipal_total": 331.16,
        "municipal_total_parcelamento": 406.92,
        "estadual_certidoes": 1,
        "analise_previa_indeferida": True,
    },
    "AMOSTRAS/AMOSTRA_2": {
        "estadual_certidoes": 2,
    },
}


def main():
    ap = argparse.ArgumentParser()
    ap.add_argument("-v", "--verbose", action="store_true", help="mostra todos os fatos, não só as diferenças")
    a = ap.parse_args()

    falhas = 0
    verificados = 0
    for caso, esperado in CASOS_ESPERADOS.items():
        caminho = RAIZ / "tests" / "fixtures" / caso
        if not caminho.is_dir():
            print(f"[PULADO] {caso}: pasta não existe ({caminho})")
            falhas += 1
            continue
        fatos = extrair_fatos(carregar_caso(caminho))
        for chave, valor_esperado in esperado.items():
            verificados += 1
            valor_obtido = fatos.get(chave)
            ok = valor_obtido == valor_esperado
            if not ok:
                falhas += 1
            if a.verbose or not ok:
                marca = "✓" if ok else "✗ QUEBROU"
                print(f"{marca} {caso}: {chave} = {valor_obtido} (esperado {valor_esperado})")

    print(f"\n{'=' * 70}\n{verificados - falhas}/{verificados} fatos conferidos batendo")
    sys.exit(1 if falhas else 0)


if __name__ == "__main__":
    main()
