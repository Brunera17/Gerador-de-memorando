#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Interface web do gerador de memorando — sem terminal, sem linha de comando.

A equipe sobe os PDFs de consulta fiscal (todos juntos, não precisa separar
por âmbito — o script classifica sozinho), confere os valores extraídos na
tela antes de gerar, ajusta os honorários se quiser, e baixa o .docx pronto.

Roda toda a lógica de extração/preenchimento de scripts/processar_fiscal.py
(processar_pasta, preencher_docx) — nada é reimplementado aqui, só a tela.

Uso local:
  pip install -r requirements.txt
  streamlit run app.py
"""
import importlib.util
import io
import sys
import tempfile
from pathlib import Path

import streamlit as st

RAIZ = Path(__file__).resolve().parent
sys.path.insert(0, str(RAIZ / "scripts"))
spec = importlib.util.spec_from_file_location("pf", RAIZ / "scripts" / "processar_fiscal.py")
pf = importlib.util.module_from_spec(spec)
spec.loader.exec_module(pf)

st.set_page_config(page_title="Gerador de Memorando", page_icon="📄", layout="wide")


# ── Acesso por senha ─────────────────────────────────────────────────────
# Só entra em vigor quando "senha_acesso" está configurada nos Secrets do
# Streamlit (é assim que se ativa no deploy na nuvem — ver .streamlit/
# secrets.toml.example). Sem essa configuração — uso local no escritório —
# a tela abre direto, sem exigir senha.
def _senha_protegida():
    try:
        return bool(st.secrets.get("senha_acesso"))
    except Exception:
        return False


def verificar_acesso():
    if not _senha_protegida():
        return True
    if st.session_state.get("autenticado"):
        return True

    st.title("📄 Gerador de Memorando")
    st.caption("Acesso restrito — dados de clientes.")
    with st.form("login"):
        senha = st.text_input("Senha de acesso", type="password")
        entrar = st.form_submit_button("Entrar")
    if entrar:
        if senha == st.secrets["senha_acesso"]:
            st.session_state["autenticado"] = True
            st.rerun()
        else:
            st.error("Senha incorreta.")
    return False


if not verificar_acesso():
    st.stop()


# ── Templates disponíveis ────────────────────────────────────────────────
# Os modelos com o timbrado do escritório ficam na raiz do projeto e NÃO são
# versionados (contêm a identidade visual do escritório) — o repositório só
# traz os modelos de templates/, sem timbrado, como ponto de partida.
# Rótulo claro (o que é + se tem timbrado) em vez do nome cru do arquivo.
def templates_disponiveis():
    candidatos = [
        ("Memorando de Consulta Fiscal — com timbrado do escritório", RAIZ / "MODELO_MEMORANDO_n8n.docx"),
        ("Memorando de Consulta Fiscal — sem timbrado (modelo genérico)", RAIZ / "templates" / "modelo_memorando.docx"),
    ]
    return {nome: caminho for nome, caminho in candidatos if caminho.is_file()}


def formatar_moeda(v):
    return pf.fmt(v) if isinstance(v, (int, float)) else str(v)


# ── Manual de uso (lido direto do MANUAL_DE_USO.md — sem duplicar texto) ──
@st.dialog("Manual de uso", width="large")
def mostrar_manual():
    caminho_manual = RAIZ / "MANUAL_DE_USO.md"
    if caminho_manual.is_file():
        st.markdown(caminho_manual.read_text(encoding="utf-8"))
    else:
        st.warning("MANUAL_DE_USO.md não encontrado no projeto.")


# ── Barra lateral ─────────────────────────────────────────────────────────
with st.sidebar:
    st.header("Configuração")

    modelos = templates_disponiveis()
    if not modelos:
        st.error("Nenhum modelo .docx encontrado na pasta do projeto.")
        st.stop()
    nome_modelo = st.selectbox("Modelo do memorando", list(modelos.keys()))
    template_path = modelos[nome_modelo]
    st.caption(f"Arquivo: `{template_path.name}`")

    st.caption(
        "Os honorários são calculados automaticamente a partir das pendências "
        "encontradas. Preencha abaixo só se quiser usar um valor fixo."
    )
    honorarios_manual = st.number_input(
        "Honorários (R$) — 0 = calcular automático", min_value=0.0, step=10.0, value=0.0
    )

    st.divider()
    st.caption("scripts/processar_fiscal.py — mesma lógica usada na linha de comando")

    st.markdown(
        """
        <style>
        div.st-key-btn_manual button {
            border-radius: 50%;
            width: 38px;
            height: 38px;
            padding: 0;
            font-family: Georgia, "Times New Roman", serif;
            font-style: italic;
            font-weight: 700;
            font-size: 16px;
            line-height: 1;
        }
        </style>
        """,
        unsafe_allow_html=True,
    )
    with st.container(key="btn_manual"):
        if st.button("i", help="Manual de uso"):
            mostrar_manual()


st.title("📄 Gerador de Memorando")
st.write(
    "Suba todos os PDFs da consulta fiscal (federal, estadual e municipal juntos — "
    "não precisa separar em pastas). O programa classifica cada um automaticamente."
)

arquivos = st.file_uploader(
    "PDFs da consulta fiscal", type=["pdf"], accept_multiple_files=True
)

if "resultado" not in st.session_state:
    st.session_state.resultado = None
    st.session_state.pasta_tmp = None

col_a, col_b = st.columns([1, 4])
processar = col_a.button("Processar PDFs", type="primary", disabled=not arquivos)

if processar and arquivos:
    with st.spinner("Lendo e classificando os PDFs..."):
        pasta_tmp = tempfile.mkdtemp(prefix="memorando_")
        for arq in arquivos:
            (Path(pasta_tmp) / arq.name).write_bytes(arq.getvalue())
        try:
            resultado = pf.processar_pasta(pasta_tmp, honorarios=0.0)
            st.session_state.resultado = resultado
            st.session_state.pasta_tmp = pasta_tmp
        except Exception as e:
            st.error(f"Não consegui processar os PDFs: {e}")
            st.session_state.resultado = None

resultado = st.session_state.resultado

if resultado:
    st.divider()
    st.subheader("Conferência antes de gerar")

    col1, col2 = st.columns(2)
    col1.metric("Empresa", resultado["nome"] or "não identificada")
    col2.metric("CNPJ", resultado["cnpj"] or "não identificado")
    if not resultado["nome"] or not resultado["cnpj"]:
        st.warning(
            "Nome ou CNPJ não foram identificados — confira se o PDF da Situação "
            "Fiscal (ou equivalente) foi enviado."
        )

    with st.expander("Classificação de cada PDF enviado", expanded=True):
        desconhecidos = [a for a, t in resultado["classificacao"] if t == "desconhecido"]
        for arq, tipo in resultado["classificacao"]:
            if tipo == "desconhecido":
                st.markdown(f"⚠️ **{arq}** → `desconhecido` (não foi usado no memorando)")
            else:
                st.markdown(f"✅ {arq} → `{tipo}`")
        if desconhecidos:
            st.info(
                "PDFs marcados como `desconhecido` não entram no memorando — é um "
                "formato que o script ainda não reconhece. Fale com quem mantém o "
                "script se isso for inesperado."
            )

    s = resultado["resumo"]
    with st.expander("Resumo por âmbito", expanded=True):
        c1, c2, c3 = st.columns(3)
        c1.markdown(f"**Federal** — {s['federal_status']}")
        c1.write(f"Parc. Simples Nacional: {s['parc_sn']}")
        c1.write(f"Parc. Simplificado: {s['parc_simplificado']}")
        c1.write(f"Parc. Dívida Ativa: {s['parc_dau']}")
        c1.write(f"Parc. SISPAR: {', '.join(s['parc_sispar']) if s['parc_sispar'] else 'não'}")
        if s["siefpar_atraso"]:
            c1.write(f"⚠️ Parcelamento em atraso: {s['siefpar_atraso']}")
        if s["analise_previa_indeferida"]:
            c1.error("⚠️ Simples Nacional será INDEFERIDO — ver Relatório de Análise Prévia")

        c2.markdown(f"**Estadual** — {s['estadual_status']}")
        if s["site_contribuinte_omissoes"]:
            c2.write(f"GIA/EFD omissas: {s['site_contribuinte_omissoes']}")

        c3.markdown(f"**Municipal** — {s['municipal_status']}")

    if honorarios_manual > 0:
        dados_finais = dict(resultado["dados_template"])
        dados_finais["RESUMO"] = pf.montar_resumo(
            resultado["dados_json"]["federal"],
            resultado["dados_json"]["declaracoes_omissas"],
            resultado["dados_json"]["municipal"],
            resultado["dados_json"]["federal"].get("parcelamento_siefpar"),
            honorarios_manual,
            pf.date.today().strftime("%d.%m.%Y"),
            resultado["dados_json"]["parcelamento_sn"],
            resultado["dados_json"]["parcelamento_simplificado"],
            resultado["dados_json"]["parcelamento_dau"],
            resultado["dados_json"]["parcelamento_sispar"],
            resultado["dados_json"]["regularize_valores"],
        )
    else:
        dados_finais = resultado["dados_template"]

    st.divider()
    st.subheader("Gerar memorando")

    with st.spinner("Preenchendo o modelo..."):
        saida_tmp = Path(tempfile.mkdtemp(prefix="memorando_out_")) / "Memorando.docx"
        pf.preencher_docx(str(template_path), str(saida_tmp), dados_finais)
        conteudo_docx = saida_tmp.read_bytes()

    nome_arquivo = f"Memorando_{(resultado['nome'] or 'sem_nome').replace(' ', '_')}.docx"
    st.download_button(
        "⬇️ Baixar memorando (.docx)",
        data=conteudo_docx,
        file_name=nome_arquivo,
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
        type="primary",
    )

    with st.expander("Dados extraídos (JSON — para conferência técnica)"):
        st.json(resultado["dados_json"])
