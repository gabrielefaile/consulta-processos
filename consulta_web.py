import streamlit as st
import requests
import re
import pandas as pd
import unicodedata
import time
from datetime import datetime
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill

# =====================================================
# CONFIGURAÇÕES
# =====================================================
st.set_page_config(page_title="Consulta de Processos - IA", layout="wide")

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.36"
    ),
    "Accept-Language": "pt-BR,pt;q=0.9",
}

TRIBUNAIS = {
    "TJSP": {
        "nome": "TJ São Paulo",
        "base_url": "https://esaj.tjsp.jus.br/cpopg",
        "search_url": "https://esaj.tjsp.jus.br/cpopg/search.do",
    },
    "TJMS": {
        "nome": "TJ Mato Grosso do Sul",
        "base_url": "https://esaj.tjms.jus.br/cpopg",
        "search_url": "https://esaj.tjms.jus.br/cpopg/search.do",
    },
    "TJMT": {
        "nome": "TJ Mato Grosso",
        "base_url": "https://consultaprocessual.tjmt.jus.br",
        "search_url": "https://consultaprocessual.tjmt.jus.br/processos",
    },
}

# =====================================================
# FUNÇÕES UTILITÁRIAS
# =====================================================
def limpar_texto(texto):
    if not texto:
        return ""
    texto = unicodedata.normalize("NFKD", texto)
    texto = texto.encode("ascii", "ignore").decode("ascii")
    return re.sub(r"\s+", " ", texto).strip()

def normalizar_valor(texto):
    if not texto:
        return None
    texto = texto.replace("R$", "").replace(".", "").replace(",", ".").strip()
    numeros = re.findall(r"[\d\.]+", texto)
    if numeros:
        try:
            return float(numeros[0])
        except ValueError:
            pass
    return None

def extrair_ano_processo(numero):
    match = re.search(r"\.(\d{4})\.", numero)
    return int(match.group(1)) if match else None

def identificar_tribunal(numero_processo):
    numero = numero_processo.replace("-", "").replace(".", "")
    if len(numero) >= 13:
        tribunal = numero[13:16]
        if tribunal == "826":
            return "TJSP"
        elif tribunal == "812":
            return "TJMS"
        elif tribunal == "811":
            return "TJMT"
    return None

# =====================================================
# FUNÇÃO IA SIMPLES E GRATUITA
# =====================================================
def ia_analisar_oportunidade(processo):
    """
    IA simples sem API: usa regras e heurísticas para classificar
    o processo como boa oportunidade ou não.
    """
    motivos = []
    score = 0

    valor = processo.get("valor") or 0
    classe = processo.get("classe", "").upper()
    assunto = processo.get("assunto", "").upper()
    status = processo.get("status", "").upper()
    movimentacao = processo.get("ultima_movimentacao", "").upper()

    # Critérios positivos
    if valor and valor >= 100000:
        score += 30
        motivos.append(f"Valor expressivo: R$ {valor:,.2f}")

    if any(p in classe for p in [
        "EXECUCAO", "COBRANCA", "MONITORIA", "RECUPERAÇÃO", "JUIZADO ESPECIAL"
    ]):
        score += 25
        motivos.append(f"Classe interessante: {classe.title()}")

    if any(p in assunto for p in [
        "CONTRATO", "FINANCIAMENTO", "HIPOTECA", "ALIENAÇÃO", "CRÉDITO"
    ]):
        score += 20
        motivos.append("Alta probabilidade de ativo financeiro")

    if any(p in movimentacao for p in [
        "HASTA PUBLICA", "LEILAO", "ARREMATAÇÃO", "PENHORA", "SEQUESTRO"
    ]):
        score += 20
        motivos.append("Existência de garantia/bem penhorado")

    # Critérios negativos
    if any(p in status for p in ["ARQUIVADO", "EXTINTO", "SUSPENSO"]):
        score -= 40
        motivos.append("Processo inativo ou arquivado")

    if any(p in classe for p in ["FAMILIA", "SEPARACAO", "ALIMENTOS", "CRIMINAL"]):
        score -= 30
        motivos.append("Classe pouco atrativa para compra de dívida")

    if valor and valor < 10000:
        score -= 15
        motivos.append("Valor baixo")

    # Classificação
    if score >= 60:
        classificacao = "🟢 ÓTIMA OPORTUNIDADE"
    elif score >= 30:
        classificacao = "🟡 OPORTUNIDADE MODERADA"
    elif score >= 0:
        classificacao = "🟠 ANALISAR COM CAUTELA"
    else:
        classificacao = "🔴 NÃO RECOMENDADO"

    return {
        "classificacao": classificacao,
        "score": score,
        "analise": "; ".join(motivos),
        "recomendacao": (
            "Recomendado prosseguir com análise detalhada." if score >= 30
            else "Verificar viabilidade jurídica antes de avançar."
        ),
        "riscos": (
            "Riscos: valor pode estar desatualizado; verificar garantias e prescrição."
            if score >= 30
            else "Riscos: processo pode estar inativo, sem garantia ou de baixo valor."
        ),
    }

# =====================================================
# SCRAPING TJSP E TJMS
# =====================================================
def buscar_tjsp_tjms(tribunal, cidade, progress=None, total=1):
    url = TRIBUNAIS[tribunal]["search_url"]
    resultados = []

    foro_map = {
        "RIBEIRAO PRETO": "0401",
        "RIBEIRÃO PRETO": "0401",
        "SAO PAULO": "0100",
        "SÃO PAULO": "0100",
        "CAMPINAS": "0300",
        "CUIABA": "0301",
        "CUIABÁ": "0301",
        "CAMPO GRANDE": "0201",
    }

    foro = foro_map.get(limpar_texto(cidade).upper(), "")
    if not foro:
        return []

    params = {
        "conversationId": "",
        "paginaConsulta": "1",
        "localPesquisa.cdLocal": foro,
        "cbPesquisa": "NUMOAB",
        "tipoNuProcesso": "UNIFICADO",
        "foroNumeroUnificado": foro,
        "dadosConsulta.valorConsulta": cidade,
        "dadosConsulta.valorConsulta2": "",
    }

    try:
        response = requests.get(url, headers=HEADERS, params=params, timeout=20)
        response.raise_for_status()
    except requests.RequestException:
        return []

    soup = BeautifulSoup(response.text, "html.parser")
    processos = soup.find_all("div", class_=re.compile("listagem-dados"))

    for proc in processos:
        try:
            numero = proc.find("span", class_="fonteNegrito") or proc.find("a", class_="linkProcesso")
            numero = numero.get_text(strip=True) if numero else ""
            partes_div = proc.find_all("td", class_="fonte")
            partes = [limpar_texto(p.get_text()) for p in partes_div]

            valor_div = proc.find("td", string=re.compile("Valor", re.I))
            valor_texto = valor_div.find_next_sibling("td").get_text(strip=True) if valor_div else ""
            valor = normalizar_valor(valor_texto)

            movimentacao = proc.find("div", id=re.compile("movimentacao"))
            ultima_mov = movimentacao.get_text(strip=True) if movimentacao else ""

            processo = {
                "tribunal": tribunal,
                "numero_processo": numero,
                "partes": " | ".join(partes),
                "valor": valor,
                "valor_formatado": f"R$ {valor:,.2f}" if valor else "Não informado",
                "classe": "",
                "assunto": "",
                "status": "Ativo",
                "ultima_movimentacao": ultima_mov,
                "ano": extrair_ano_processo(numero),
                "cidade": cidade,
                "link": response.url,
            }
            processo["ia"] = ia_analisar_oportunidade(processo)
            resultados.append(processo)
        except Exception:
            continue

    return resultados

# =====================================================
# TJMT (nova consulta processual)
# =====================================================
def buscar_tjmt(cidade, progress=None, total=1):
    url = TRIBUNAIS["TJMT"]["search_url"]
    resultados = []

    payload = {
        "pagina": 1,
        "tamanhoPagina": 50,
        "tipoConsulta": "parte",
        "termo": cidade,
    }

    try:
        response = requests.post(url, headers=HEADERS, json=payload, timeout=20)
        response.raise_for_status()
    except requests.RequestException:
        return []

    try:
        data = response.json()
    except ValueError:
        return []

    for item in data.get("processos", []):
        numero = item.get("numeroProcesso", "")
        partes = item.get("partes", [])
        partes_texto = " | ".join([p.get("nome", "") for p in partes if p.get("nome")])

        processo = {
            "tribunal": "TJMT",
            "numero_processo": numero,
            "partes": partes_texto,
            "valor": normalizar_valor(item.get("valorCausa")),
            "valor_formatado": item.get("valorCausa", "Não informado"),
            "classe": item.get("classeProcessual", ""),
            "assunto": item.get("assunto", ""),
            "status": item.get("situacao", "Ativo"),
            "ultima_movimentacao": item.get("ultimaMovimentacao", ""),
            "ano": extrair_ano_processo(numero),
            "cidade": cidade,
            "link": f"https://consultaprocessual.tjmt.jus.br/detalhe/{numero}",
        }
        processo["ia"] = ia_analisar_oportunidade(processo)
        resultados.append(processo)

    return resultados

# =====================================================
# GERAR EXCEL
# =====================================================
def gerar_excel(processos, cidade):
    df = pd.DataFrame([
        {
            "Tribunal": p["tribunal"],
            "Número do Processo": p["numero_processo"],
            "Ano": p["ano"],
            "Cidade": p["cidade"],
            "Classe": p["classe"],
            "Assunto": p["assunto"],
            "Status": p["status"],
            "Valor": p["valor_formatado"],
            "Última Movimentação": p["ultima_movimentacao"],
            "Partes": p["partes"],
            "Link": p["link"],
            "Classificação IA": p["ia"]["classificacao"],
            "Score IA": p["ia"]["score"],
            "Análise IA": p["ia"]["analise"],
            "Recomendação IA": p["ia"]["recomendacao"],
            "Riscos IA": p["ia"]["riscos"],
        }
        for p in processos
    ])

    nome_arquivo = f"processos_{limpar_texto(cidade).lower().replace(' ', '_')}_{datetime.now().strftime('%d%m%Y')}.xlsx"
    caminho = f"C:/Users/User/ConsultaProcessos/{nome_arquivo}"
    df.to_excel(caminho, index=False, engine="openpyxl")

    # Estilizar
    wb = Workbook()
    wb = pd.read_excel(caminho, sheet_name=0)
    # (Aqui você pode adicionar cores nas colunas depois)

    return caminho, df

# =====================================================
# INTERFACE STREAMLIT
# =====================================================
def main():
    st.title("🤖 Consulta de Processos com IA")
    st.markdown("Busque processos por cidade nos tribunais **TJSP**, **TJMS** e **TJMT**.")

    with st.form("form_busca"):
        col1, col2 = st.columns([3, 1])
        with col1:
            cidade = st.text_input("Cidade", placeholder="Ex: Ribeirão Preto")
        with col2:
            st.markdown("<br>", unsafe_allow_html=True)
            tribunal_opcao = st.selectbox(
                "Tribunal",
                ["Todos", "TJSP", "TJMS", "TJMT"]
            )
        submit = st.form_submit_button("🔍 Buscar Processos")

    if submit and cidade:
        st.info(f"Buscando processos em **{cidade}**...")

        tribunais_buscar = []
        if tribunal_opcao == "Todos":
            tribunais_buscar = ["TJSP", "TJMS", "TJMT"]
        else:
            tribunais_buscar = [tribunal_opcao]

        todos_processos = []
        progresso = st.progress(0)
        status_text = st.empty()

        for i, tribunal in enumerate(tribunais_buscar):
            status_text.text(f"Buscando no {TRIBUNAIS[tribunal]['nome']}...")
            if tribunal == "TJMT":
                processos = buscar_tjmt(cidade)
            else:
                processos = buscar_tjsp_tjms(tribunal, cidade)
            todos_processos.extend(processos)
            progresso.progress((i + 1) / len(tribunais_buscar))
            time.sleep(0.5)

        progresso.empty()
        status_text.empty()

        if not todos_processos:
            st.warning("Nenhum processo encontrado. Tente outra cidade ou tribunal.")
            return

        st.success(f"{len(todos_processos)} processo(s) encontrado(s).")

        # Exibir resultados
        for p in todos_processos:
            with st.expander(f"{p['tribunal']} - {p['numero_processo']} - {p['ia']['classificacao']}"):
                col1, col2 = st.columns(2)
                with col1:
                    st.write(f"**Valor:** {p['valor_formatado']}")
                    st.write(f"**Classe:** {p['classe'] or 'Não informado'}")
                    st.write(f"**Status:** {p['status']}")
                with col2:
                    st.write(f"**Ano:** {p['ano'] or 'Não informado'}")
                    st.write(f"**Assunto:** {p['assunto'] or 'Não informado'}")
                    st.write(f"**Última Movimentação:** {p['ultima_movimentacao'] or 'Não informado'}")
                st.write(f"**Partes:** {p['partes'] or 'Não informado'}")
                st.write(f"**Link:** [{p['link']}]({p['link']})")
                st.markdown("---")
                st.write(f"**Score IA:** {p['ia']['score']}")
                st.write(f"**Análise:** {p['ia']['analise']}")
                st.write(f"**Recomendação:** {p['ia']['recomendacao']}")
                st.write(f"**Riscos:** {p['ia']['riscos']}")

        # Gerar Excel
        caminho, df = gerar_excel(todos_processos, cidade)
        st.download_button(
            label="📥 Baixar planilha Excel",
            data=open(caminho, "rb"),
            file_name=caminho.split("/")[-1],
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )

        st.info(f"Planilha também salva em: `{caminho}`")

if __name__ == "__main__":
    main()
