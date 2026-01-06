from selenium import webdriver
from selenium.webdriver.support.ui import Select
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.webdriver import ActionChains
from selenium.webdriver.common.keys import Keys
import pandas as pd
from bs4 import BeautifulSoup
import time
from datetime import datetime, timedelta
from datetime import date
import pandas as pd
import unicodedata
import re

driver = webdriver.Firefox()
driver.get('https://fnet.bmfbovespa.com.br/fnet/publico/abrirGerenciadorDocumentosCVM')
time.sleep(1)

filtro = driver.find_element(value = 'showFiltros')
filtro.click()
time.sleep(1)

hoje = date.today()
ontem = hoje - timedelta(days=1)
valor = f"{ontem.day:02d}/{ontem.month:02d}/{ontem.year}"

el = WebDriverWait(driver, 10).until(
    EC.element_to_be_clickable((By.ID, "dataInicial"))
)

el.clear()
el.send_keys(valor)

def select2_by_text_click(driver, container_id, option_text, timeout=10):
    # 1) abrir o select2
    container = WebDriverWait(driver, timeout).until(
        EC.element_to_be_clickable((By.ID, container_id))
    )
    container.click()

    # 2) esperar o dropdown aparecer (v3 ou v4)
    # v3 usa 'select2-drop' e 'select2-result-label'
    # v4 usa 'select2-dropdown' e 'select2-results__option'
    try:
        # tenta v3
        results_label = WebDriverWait(driver, timeout).until(
            EC.visibility_of_element_located((By.XPATH,
                f"//li[contains(@class,'select2-result') or contains(@class,'select2-result-selectable')]//div[contains(@class,'select2-result-label') and normalize-space(.)='{option_text}']"))
        )
        results_label.click()
        return True
    except Exception:
        pass

    try:
        # tenta v4
        option = WebDriverWait(driver, timeout).until(
            EC.visibility_of_element_located((By.XPATH,
                f"//li[contains(@class,'select2-results__option') and normalize-space(.)='{option_text}']"))
        )
        option.click()
        return True
    except Exception:
        pass

    return False

select2_by_text_click(driver, "s2id_tipoFundo", "Fundo Imobiliário")
filtrar = driver.find_element(value = "filtrar")
filtrar.click()
time.sleep(5)

dropdown = driver.find_element(By.CSS_SELECTOR, 'div#tblDocumentosEnviados_length select')

# Create a Select object
select = Select(dropdown)

# Select the "100" option by value
select.select_by_value('100')
time.sleep(1)

def read_table():

    table_element = driver.find_element(value = "tblDocumentosEnviados")

    # Get the HTML content of the table
    table_html = table_element.get_attribute('outerHTML')

    # Parse the HTML content using BeautifulSoup
    soup = BeautifulSoup(table_html, 'html.parser')

    # Extract the table data using BeautifulSoup
    table_data = []
    for row in soup.find_all('tr'):
        row_data = []
        for cell in row.find_all(['th', 'td']):
            if cell.find('a'):  # Check if the cell contains a link
                link = cell.find('a')
                cell_data = link['href'] #{
                    #'text': link.text.strip(),
                    #'url': link['href']
                #}
            else:
                cell_data = cell.text.strip()
            row_data.append(cell_data)
        table_data.append(row_data)

    # Convert the table data into a Pandas DataFrame
    df = pd.DataFrame(table_data[1:], columns=table_data[0])
    df['DocNumber'] = df['Ações'].apply(lambda x: x.replace('visualizarDocumento?id=', '').replace('&cvm=true', ''))
    df = df.drop(columns='Ações')

    return df

next = driver.find_element(value = "tblDocumentosEnviados_next")

pages = []
while next.get_attribute('class') != 'paginate_button next disabled':
    pages.append(read_table())
    next.click()
    time.sleep(10)
    next = driver.find_element(value = "tblDocumentosEnviados_next")

pages.append(read_table())

driver.close()

data = (
    pd.concat(pages)
    .rename(columns = {'Nome do Fundo': 'Nome_Fundo', 'Data de Referência': 'Dt_Ref', 'Data de Entrega': 'Dt_Entrega', 'Espécie': 'Especie'})
    .assign(Dt_Entrega = lambda x: pd.to_datetime(x['Dt_Entrega'], dayfirst=True))
    .assign(Dt_Ref = lambda x: pd.to_datetime(x['Dt_Ref'], format='mixed', dayfirst=True))
)

nomes_raw = [
    "VALORA CRI ÍNDICE DE PREÇO FUNDO DE INVESTIMENTO IMOBILIÁRIO - FII RESPONSABILIDADE LIMITADA",
    "CAPITÂNIA SECURITIES II FUNDO DE INVESTIMENTO IMOBILIÁRIO RESPONSABILIDADE LIMITADA",
    "XP SELECTION FUNDO DE FUNDOS DE INVESTIMENTO IMOBILIÁRIO - FII",
    "RBR ALPHA MULTIESTRATÉGIA REAL ESTATE FUNDO DE INVESTIMENTO IMOBILIÁRIO RESPONSABILIDADE LIMITADA",
    "HEDGE TOP FOFII 3 FUNDO DE INVESTIMENTO IMOBILIÁRIO DE RESPONSABILIDADE LIMITADA",
    "FUNDO DE INVESTIMENTO IMOBILIÁRIO - FII REC RENDA IMOBILIÁRIA - RESPONSABILIDADE LIMITADA",
    "BTG PACTUAL CORPORATE OFFICE FUND - FUNDO DE INVESTIMENTO IMOBILIÁRIO RESPONSABILIDADE LIMITADA",
    "BTG PACTUAL SHOPPINGS FUNDO DE INVESTIMENTO IMOBILIÁRIO RESPONSABILIDADE LIMITADA",
    "BTG PACTUAL AGRO LOGÍSTICA FUNDO DE INVESTIMENTO IMOBILIÁRIO RESPONSABILIDADE LIMITADA",
    "FUNDO DE INVESTIMENTO IMOBILIÁRIO REC LOGÍSTICA - RESPONSABILIDADE LIMITADA"
]

def normalize(s):
    if pd.isna(s):
        return ""
    s = str(s).strip().casefold()                    # trim + casefold
    s = unicodedata.normalize("NFD", s)              # decompor acentos
    s = re.sub(r"[\u0300-\u036f]", "", s)            # remover marcadores de acento
    s = re.sub(r"\s+", " ", s)                       # normalizar espaços múltiplos
    return s

# --- carregar / ter o DataFrame (ex.: df já existe). Exemplo de leitura:
# df = pd.read_csv("tabela.csv")   # ou df = seu_dataframe
df = data
# normalizar coluna e a lista de nomes
df["nome_norm"] = df["Nome_Fundo"].apply(normalize)
target_set = {normalize(x) for x in nomes_raw}      # set para lookup rápido

# --- FILTRO EXATO (recomendado quando quer correspondência inteira) ---
filtered = df[df["nome_norm"].isin(target_set)].copy()

# --- opcional: FILTRO POR SUBSTRING (se quiser manter linhas que contenham qualquer dos nomes) ---
# construir regex seguro (escape) a partir da lista normalizada
pattern = "|".join(re.escape(x) for x in target_set)
filtered_contains = df[df["nome_norm"].str.contains(pattern, na=False)].copy()

# --- exportar ou usar ---
filtered.to_csv("filtrados_exatos.csv", index=False)
# filtered_contains.to_csv("filtrados_contains.csv", index=False)

# print("exatos:", len(filtered), "contendo (substring):", len(filtered_contains))

frases = [
    "Fato Relevante",
    "Aviso aos Cotistas - Estruturado"
]

# --- normaliza as frases target e cria regex (escaped) ---
frases_norm = [normalize(f) for f in frases]
pattern = "|".join(re.escape(f) for f in frases_norm)  # regex: frase1|frase2

# --- 1) tentar identificar coluna de categoria automaticamente ---
def guess_category_column(df):
    candidates = []
    for col in df.columns:
        col_norm = normalize(col)
        # procurar termos comuns em nomes de colunas
        if any(k in col_norm for k in ("Categoria", "categoria")):
            candidates.append(col)
    return candidates[0] if candidates else None

cat_col = guess_category_column(filtered)

if cat_col:
    # cria uma coluna normalizada da categoria
    filtered["cat_norm"] = filtered[cat_col].apply(normalize)
    # 2 opções: exact match (isin) ou contains (substring). Aqui usamos contains (mais flexível)
    result = filtered[filtered["cat_norm"].str.contains(pattern, na=False)].copy()
else:
    # fallback: procurar nas colunas textuais (concatena conteúdo das colunas string e busca)
    text_cols = [c for c in filtered.columns if filtered[c].dtype == "object"]
    if not text_cols:
        raise ValueError("Nenhuma coluna textual encontrada para procurar as frases.")
    # cria coluna temporária com concat das colunas textuais normalizadas
    def concat_norm(row):
        parts = [normalize(row[c]) for c in text_cols]
        return " ".join(p for p in parts if p)
    filtered["all_text_norm"] = filtered.apply(concat_norm, axis=1)
    result = filtered[filtered["all_text_norm"].str.contains(pattern, na=False)].copy()

# --- opcional: remover colunas norm temporárias ou manter para debug ---
# result = result.drop(columns=[c for c in ("cat_norm", "all_text_norm") if c in result.columns])

# --- salvar ou usar ---

chaves = [
    "AGE",
    "Relatório Gerencial"
]

chaves_norm = [normalize(x) for x in chaves]
pattern = "|".join(re.escape(x) for x in chaves_norm)  # ex: "age|relatorio gerencial"

# --- escolher DataFrame base: tenta usar 'result' (do passo anterior), senão 'filtered', senão 'df' ---
try:
    base_df = filtered
except NameError:
        base_df = filtered


# --- tentar adivinhar o nome da coluna 'tipo' (ou setar manualmente abaixo) ---
def guess_type_column(df):
    candidates = []
    for col in df.columns:
        col_norm = normalize(col)
        if any(k in col_norm for k in ("tipo", "Tipo")):
            candidates.append(col)
    return candidates[0] if candidates else None

tipo_col = guess_type_column(base_df)
# --- se souber o nome exato da coluna, descomente e ajuste:
# tipo_col = "Tipo"  # ou "tipo", "TipoDocumento", etc.

if tipo_col:
    base_df["tipo_norm"] = base_df[tipo_col].apply(normalize)
    final = base_df[base_df["tipo_norm"].str.contains(pattern, na=False)].copy()
else:
    # fallback: procurar em todas colunas string (concatena e busca)
    text_cols = [c for c in base_df.columns if base_df[c].dtype == "object"]
    if not text_cols:
        raise ValueError("Nenhuma coluna textual encontrada para procurar as chaves.")
    def concat_norm(row):
        parts = [normalize(row[c]) for c in text_cols]
        return " ".join(p for p in parts if p)
    base_df["all_text_norm"] = base_df.apply(concat_norm, axis=1)
    final = base_df[base_df["all_text_norm"].str.contains(pattern, na=False)].copy()

# --- opcional: remover colunas temporárias criadas ---
for tmp in ("tipo_norm", "all_text_norm"):
    if tmp in final.columns:
        # se quiser manter para debug, comente a linha abaixo
        final = final.drop(columns=[tmp])

# --- resultado ---

dfFinal = pd.concat([result,final])

import yagmail, os

USER = "viniciuskaolo@gmail.com"           # seu gmail
APP_PASS = "hjlv knog yxjt muku"   # app password (16 chars)

links = []

for row in dfFinal.itertuples(index=False):  # index=False para não incluir Index no tuple
    valor = getattr(row, "DocNumber")    # ou row.DocNumber
    if pd.isna(valor):
        continue

    # pegar Categoria e Tipo (tratando NaNs)
    cat = getattr(row, "Categoria")
    tipo = getattr(row, "Tipo")

    parts = []
    if pd.notna(cat):
        parts.append(str(cat).strip())
    if pd.notna(tipo):
        parts.append(str(tipo).strip())

    prefix = " - ".join(parts) if parts else "Sem categoria/tipo"

    link = f"https://fnet.bmfbovespa.com.br/fnet/publico/visualizarDocumento?id={valor}&cvm=true"
    links.append(f"{prefix}: {link}")

# --- fallback em texto (bom para clientes que não renderizam HTML) ---
text_body = "Olá,\n\nSeguem os links dos documentos:\n\n" + "\n".join(links)

# --- versão HTML com lista clicável ---
# (pode personalizar o texto do link, ex: "Documento {i+1}" se preferir)
# --- enviar com yagmail (texto + HTML) ---
yag = yagmail.SMTP(USER, APP_PASS)
yag.send(
    to=["diego@seikopartners.com.br","fabio@seikopartners.com.br"],
    subject="Relatório dos Fundo Imobiliários",
    contents=[text_body],  # primeiro texto, depois html
)
yag.close()
print("Email enviado com", len(links), "links.")