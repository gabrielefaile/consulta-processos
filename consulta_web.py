import streamlit as st
import requests
import re
import pandas as pd
import unicodedata
import io
import time
from datetime import datetime
from bs4 import BeautifulSoup
from openpyxl import Workbook
from openpyxl.styles import Font, PatternFill, Alignment

# =====================================================
# CONFIGURAÇÕES
# =====================================================
st.set_page_config(page_title="Consulta de Processos - IA", layout="wide")

HEADERS = {
    "User-Agent": (
        "Mozilla/5.0 (Windows NT 10.0; Win64; x64) "
        "AppleWebKit/537.0.36 (KHTML, like Gecko) "
        "Chrome/120.0.0.0 Safari/537.0.36"
    ),
    "Accept-Language": "pt-BR,pt;q=0.9",
    "Accept": "text/html,application/xhtml+xml,application/xml;q=0.9,*/*;q=0.8",
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
        "consulta_url": "https://consultaprocessual.tjmt.jus.br/",
    },
}

# Códigos de comarcas principais
COMARCAS = {
    "TJSP": {
        "SAO PAULO": "0100",
        "SÃO PAULO": "0100",
        "CAMPINAS": "0300",
        "RIBEIRAO PRETO": "0506",
        "RIBEIRÃO PRETO": "0506",
        "SAO JOSE DO RIO PRETO": "0618",
        "SÃO JOSÉ DO RIO PRETO": "0618",
        "SOROCABA": "0670",
        "SANTOS": "0636",
        "SAO BERNARDO DO CAMPO": "0640",
        "SÃO BERNARDO DO CAMPO": "0640",
    },
    "TJMS": {
        "CAMPO GRANDE": "0201",
        "CORUMBA": "0401",
        "CORUMBÁ": "0401",
        "DOURADOS": "0301",
        "TRES LAGOAS": "0501",
        "TRÊS LAGOAS": "0501",
    },
}

TIPOS_BUSCA = {
    "Nome da parte": "NMPARTE",
    "CPF/CNPJ da parte": "DOCPARTE",
    "OAB do advogado": "NUMOAB",
    "Número do processo": "NUMPROC",
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

def extrair_foro_processo(numero):
    match = re.search(r"\.(\d{4})$", numero.replace("-", ""))
    return match.group(1) if match else ""

# =====================================================
# IA SIMPLES
# =====================================================
def ia_analisar_oportunidade(processo):
    motivos = []
    score = 0

    valor = processo.get("valor") or 0
    classe = processo.get("classe", "").upper()
    assunto = processo.get("assunto", "").upper()
    status = processo.get("status", "").upper()
    movimentacao = processo.get("ultima_movimentacao", "").upper()

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

    if any(p in status for p in ["ARQUIVADO", "EXTINTO", "SUSPENSO"]):
        score -= 40
        motivos.append("Processo inativo ou arquivado")

    if any(p in classe for p in ["FAMILIA", "SEPARACAO", "ALIMENTOS", "CRIMINAL"]):
        score -= 30
        motivos.append("Classe pouco atrativa para compra de dívida")

    if valor and valor < 10000:
        score -= 15
        motivos.append("Valor baixo")

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
# SCRAPING E-SAJ (TJSP E TJMS)
# =====================================================
def buscar_esaj(tribunal, comarca, tipo_busca, termo):
    base = TRIBUNAIS[tribunal]["base_url"]
    cd_local = COMARCAS.get(tribunal, {}).get(limpar_texto(comarca).upper())

    if not cd_local:
        st.error(f"Comarca '{comarca}' não cadastrada para {TRIBUNAIS[tribunal]['nome']}.")
        return []

    resultados = []
    session = requests.Session()
    session.headers.update(HEADERS)

    # Passo 1: abrir página para criar sessão
    try:
        open_resp = session.get(f"{base}/open.do", timeout=20)
        open_resp.raise_for_status()
    except requests.RequestException as e:
        st.error(f"Erro ao conectar no {TRIBUNAIS[tribunal]['nome']}: {e}")
        return []

    # Passo 2: montar parâmetros de busca
    params = {
        "conversationId": "",
        "paginaConsulta": "1",
        "dadosConsulta.localPesquisa.cdLocal": cd_local,
        "cbPesquisa": tipo_busca,
        "dadosConsulta.tipoNuProcesso": "UNIFICADO",
        "dadosConsulta.valorConsulta": termo,
    }

    # Se for número do processo, usa parâmetros específicos
    if tipo_busca == "NUMPROC":
        params["dadosConsulta.valorConsultaNuUnificado"] = termo
        params["foroNumeroUnificado"] = extrair_foro_processo(termo)
        params["dadosConsulta.valorConsulta"] = ""

    # Passo 3: fazer busca
    try:
        resp = session.get(f"{base}/search.do", params=params, timeout=30)
        resp.raise_for_status()
    except requests.RequestException as e:
        st.error(f"Erro na busca no {TRIBUNAIS[tribunal]['nome']}: {e}")
        return []

    # Verifica se houve aviso de "muitos resultados" ou "nenhum"
    if "não existem informações disponíveis" in resp.text.lower():
        return []
    if "foram encontrados muitos processos" in resp.text.lower():
        st.warning("Foram encontrados muitos processos. Refine a busca.")

    soup = BeautifulSoup(resp.text, "html.parser")

    # --- tenta achar lista de resultados ---
    # O e-SAJ pode ter tabela com links de processo
    links = soup.find_all("a", class_="linkProcesso")
    if not links:
        links = soup.select("a[href*='show.do?processo.codigo=']")

    for link in links:
        numero = limpar_texto(link.get_text())
        href = link.get("href", "")
        if href.startswith("/"):
            href = f"https://{base.split('/')[2]}{href}"

        processo = {
            "tribunal": tribunal,
            "numero_processo": numero,
            "link": href,
        }
        resultados.append(processo)

    # Se não achou links, tenta extrair da estrutura de listagem
    if not resultados:
        processos_div = soup.find_all("div", class_=re.compile("listagem-dados|snippet"))
        for proc in processos_div:
            try:
                numero_tag = (
                    proc.find("a", class_="linkProcesso") or
                    proc.find("span", class_="fonteNegrito") or
                    proc.find("a")
                )
                numero = limpar_texto(numero_tag.get_text()) if numero_tag else ""
                if not numero:
                    continue

                processo = {
                    "tribunal": tribunal,
                    "numero_processo": numero,
                    "link": TRIBUNAIS[tribunal]["base_url"],
                }
                resultados.append(processo)
            except Exception:
                continue

    # Para cada resultado, abre página do processo para pegar detalhes
    processos_completos = []
    for proc in resultados[:10]:  # limita a 10 para não sobrecarregar
        detalhes = buscar_detalhes_processo(tribunal, proc["numero_processo"], proc["link"])
        processos_completos.append(detalhes)
        time.sleep(0.3)

    return processos_completos


def buscar_detalhes_processo(tribunal, numero, link):
    """Extrai detalhes da página do processo."""
    processo = {
        "tribunal": tribunal,
        "numero_processo": numero,
        "link": link,
        "partes": "",
        "valor": None,
        "valor_formatado": "Não informado",
        "classe": "",
        "assunto": "",
        "status": "Ativo",
        "ultima_movimentacao": "",
        "ano": extrair_ano_processo(numero),
        "cidade": "",
    }

    try:
        resp = requests.get(link, headers=HEADERS, timeout=20)
        resp.raise_for_status()
    except Exception:
        processo["ia"] = ia_analisar_oportunidade(processo)
        return processo

    soup = BeautifulSoup(resp.text, "html.parser")

    # Classe
    classe_label = soup.find("span", string=re.compile("Classe", re.I))
    if classe_label:
        processo["classe"] = limpar_texto(classe_label.find_next_sibling().get_text()) if classe_label.find_next_sibling() else ""

    # Assunto
    assunto_label = soup.find("span", string=re.compile("Assunto", re.I))
    if assunto_label:
        prox = assunto_label.find_next_sibling()
        processo["assunto"] = limpar_texto(prox.get_text()) if prox else ""

    # Valor
    valor_match = re.search(r"Valor da ação.*?\R?\s*([\d\.,]+)", resp.text)
    if valor_match:
        processo["valor"] = normalizar_valor(valor_match.group(1))
        processo["valor_formatado"] = f"R$ {processo['valor']:,.2f}" if processo["valor"] else "Não informado"

    # Status
    if "Baixado" in resp.text or "Extinto" in resp.text:
        processo["status"] = "Baixado/Extinto"
    elif "Arquivado" in resp.text:
        processo["status"] = "Arquivado"
    elif "Em tramitação" in resp.text:
        processo["status"] = "Ativo"

    # Partes
    partes = []
    tabelas = soup.find_all("table", class_=re.compile("table|secao"))
    for tabela in tabelas:
        for tr in tabela.find_all("tr"):
            texto = limpar_texto(tr.get_text())
            if texto and len(texto) > 5:
                partes.append(texto)
    processo["partes"] = " | ".join(partes[:5])

    # Última movimentação
    movs = soup.find_all("tr", class_=re.compile("movimentacao"))
    if movs:
        processo["ultima_movimentacao"] = limpar_texto(movs[0].get_text())

    processo["ia"] = ia_analisar_oportunidade(processo)
    return processo


# =====================================================
# TJMT - link direto (não permite scraping fácil)
# =====================================================
def buscar_tjmt(comarca, termo):
    st.info("TJMT: Consulta automatizada não disponível. Acesse o portal oficial.")
    return []


# =====================================================
# GERAR EXCEL EM MEMÓRIA
# =====================================================
def gerar_excel(processos, cidade):
    df = pd.DataFrame([
        {
            "Tribunal": p["tribunal"],
            "Número do Processo": p["numero_processo"],
            "Ano": p["ano"],
            "Cidade": p.get("cidade", cidade),
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

    buffer = io.BytesIO()
    with pd.ExcelWriter(buffer, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Processos")

    return buffer.getvalue(), df


# =====================================================
# INTERFACE STREAMLIT
# =====================================================
def main():
    st.title("🤖 Consulta de Processos com IA")
    st.markdown("Busque processos por **nome da parte**, **CPF/CNPJ**, **OAB** ou **número do processo**.")

    with st.form("form_busca"):
        col1, col2 = st.columns([2, 1])

        with col1:
            tribunal_opcao = st.selectbox("Tribunal", ["TJSP", "TJMS", "TJMT"])

        with col2:
            tipo_busca = st.selectbox("Tipo de busca", list(TIPOS_BUSCA.keys()))

        comarca = st.text_input("Comarca/Cidade", placeholder="Ex: Ribeirão Preto")
        termo = st.text_input("Termo de busca", placeholder="Ex: nome da parte, CPF, OAB ou nº do processo")

        submit = st.form_submit_button("🔍 Buscar Processos")

    if submit:
        if not comarca or not termo:
            st.warning("Preencha a comarca e o termo de busca.")
            return

        st.info(f"Buscando em **{comarca}** no **{TRIBUNAIS[tribunal_opcao]['nome']}**...")

        if tribunal_opcao == "TJMT":
            st.warning(
                "O TJMT não permite consulta automatizada direta. "
                f"Clique para consultar manualmente: [{TRIBUNAIS['TJMT']['consulta_url']}]({TRIBUNAIS['TJMT']['consulta_url']})"
            )
            return

        processos = buscar_esaj(tribunal_opcao, comarca, TIPOS_BUSCA[tipo_busca], termo)

        if not processos:
            st.warning("Nenhum processo encontrado. Verifique a comarca e o termo de busca.")
            return

        st.success(f"{len(processos)} processo(s) encontrado(s).")

        for p in processos:
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

        excel_bytes, df = gerar_excel(processos, comarca)
        st.download_button(
            label="📥 Baixar planilha Excel",
            data=excel_bytes,
            file_name=f"processos_{limpar_texto(comarca).lower().replace(' ', '_')}_{datetime.now().strftime('%d%m%Y')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )


if __name__ == "__main__":
    main()
