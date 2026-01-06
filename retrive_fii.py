from __future__ import annotations

import os
import re
import time
import unicodedata
from datetime import date, timedelta

import pandas as pd
import yagmail
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC


URL = "https://fnet.bmfbovespa.com.br/fnet/publico/abrirGerenciadorDocumentosCVM"

NOMES_RAW = [
    "VALORA CRI ÍNDICE DE PREÇO FUNDO DE INVESTIMENTO IMOBILIÁRIO - FII RESPONSABILIDADE LIMITADA",
    "CAPITÂNIA SECURITIES II FUNDO DE INVESTIMENTO IMOBILIÁRIO RESPONSABILIDADE LIMITADA",
    "XP SELECTION FUNDO DE FUNDOS DE INVESTIMENTO IMOBILIÁRIO - FII",
    "RBR ALPHA MULTIESTRATÉGIA REAL ESTATE FUNDO DE INVESTIMENTO IMOBILIÁRIO RESPONSABILIDADE LIMITADA",
    "HEDGE TOP FOFII 3 FUNDO DE INVESTIMENTO IMOBILIÁRIO DE RESPONSABILIDADE LIMITADA",
    "FUNDO DE INVESTIMENTO IMOBILIÁRIO - FII REC RENDA IMOBILIÁRIA - RESPONSABILIDADE LIMITADA",
    "BTG PACTUAL CORPORATE OFFICE FUND - FUNDO DE INVESTIMENTO IMOBILIÁRIO RESPONSABILIDADE LIMITADA",
    "BTG PACTUAL SHOPPINGS FUNDO DE INVESTIMENTO IMOBILIÁRIO RESPONSABILIDADE LIMITADA",
    "BTG PACTUAL AGRO LOGÍSTICA FUNDO DE INVESTIMENTO IMOBILIÁRIO RESPONSABILIDADE LIMITADA",
    "FUNDO DE INVESTIMENTO IMOBILIÁRIO REC LOGÍSTICA - RESPONSABILIDADE LIMITADA",
]

FRASES_CATEGORIA = [
    "Fato Relevante",
    "Aviso aos Cotistas - Estruturado",
]

CHAVES_TIPO = [
    "AGE",
    "Relatório Gerencial",
]

EMAIL_TO = ["diego@seikopartners.com.br", "fabio@seikopartners.com.br"]
EMAIL_SUBJECT = "Relatório dos Fundo Imobiliários"


def normalize(s) -> str:
    if s is None or (isinstance(s, float) and pd.isna(s)):
        return ""
    s = str(s).strip().casefold()
    s = unicodedata.normalize("NFD", s)
    s = re.sub(r"[\u0300-\u036f]", "", s)  # remove acentos
    s = re.sub(r"\s+", " ", s)
    return s


def build_contains_pattern(terms: list[str]) -> str:
    terms_norm = [normalize(t) for t in terms if t and str(t).strip()]
    terms_norm = [t for t in terms_norm if t]
    if not terms_norm:
        return ""
    return "|".join(re.escape(t) for t in terms_norm)


def guess_column(df: pd.DataFrame, keywords: list[str]) -> str | None:
    # tenta achar uma coluna cujo nome contenha alguma keyword (normalizada)
    cols = list(df.columns)
    for col in cols:
        col_norm = normalize(col)
        if any(normalize(k) in col_norm for k in keywords):
            return col
    return None


def select2_by_text_click(driver, container_id: str, option_text: str, timeout: int = 15) -> bool:
    container = WebDriverWait(driver, timeout).until(
        EC.element_to_be_clickable((By.ID, container_id))
    )
    container.click()

    # tenta select2 v3
    try:
        results_label = WebDriverWait(driver, timeout).until(
            EC.visibility_of_element_located(
                (By.XPATH,
                 f"//div[contains(@class,'select2-result-label') and normalize-space(.)='{option_text}']")
            )
        )
        results_label.click()
        return True
    except Exception:
        pass

    # tenta select2 v4
    try:
        option = WebDriverWait(driver, timeout).until(
            EC.visibility_of_element_located(
                (By.XPATH,
                 f"//li[contains(@class,'select2-results__option') and normalize-space(.)='{option_text}']")
            )
        )
        option.click()
        return True
    except Exception:
        return False


def make_driver() -> webdriver.Chrome:
    opts = webdriver.ChromeOptions()
    # headless novo (Chrome >= 109). Se der erro, troque por "--headless"
    opts.add_argument("--headless=new")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--window-size=1920,1080")
    # Em Debian/Ubuntu com chromium instalado via apt, geralmente funciona sem setar binary_location.
    # Se precisar, descomente e ajuste:
    # opts.binary_location = "/usr/bin/chromium"

    return webdriver.Chrome(options=opts)


def wait_table_ready(driver, timeout: int = 20):
    WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((By.ID, "tblDocumentosEnviados"))
    )
    # garante que o corpo da tabela carregou
    WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((By.CSS_SELECTOR, "#tblDocumentosEnviados tbody tr"))
    )


def read_table(driver) -> pd.DataFrame:
    table_element = driver.find_element(By.ID, "tblDocumentosEnviados")
    table_html = table_element.get_attribute("outerHTML")

    soup = BeautifulSoup(table_html, "html.parser")

    table_data = []
    for row in soup.find_all("tr"):
        row_data = []
        for cell in row.find_all(["th", "td"]):
            if cell.find("a"):
                link = cell.find("a")
                cell_data = link.get("href", "").strip()
            else:
                cell_data = cell.get_text(strip=True)
            row_data.append(cell_data)
        table_data.append(row_data)

    if not table_data or len(table_data) < 2:
        return pd.DataFrame()

    df = pd.DataFrame(table_data[1:], columns=table_data[0])

    # coluna "Ações" geralmente contém visualizarDocumento?id=XXXX&cvm=true
    if "Ações" in df.columns:
        df["DocNumber"] = (
            df["Ações"]
            .astype(str)
            .str.replace("visualizarDocumento?id=", "", regex=False)
            .str.replace("&cvm=true", "", regex=False)
        )
        df = df.drop(columns=["Ações"])

    return df


def collect_pages(driver) -> pd.DataFrame:
    pages = []

    wait_table_ready(driver)
    pages.append(read_table(driver))

    while True:
        next_btn = driver.find_element(By.ID, "tblDocumentosEnviados_next")
        classes = next_btn.get_attribute("class") or ""
        if "disabled" in classes:
            break

        next_btn.click()
        wait_table_ready(driver)
        pages.append(read_table(driver))

    if not pages:
        return pd.DataFrame()
    return pd.concat(pages, ignore_index=True)


def filter_df(df: pd.DataFrame) -> pd.DataFrame:
    # renomear colunas se existirem
    rename_map = {
        "Nome do Fundo": "Nome_Fundo",
        "Data de Referência": "Dt_Ref",
        "Data de Entrega": "Dt_Entrega",
        "Espécie": "Especie",
    }
    for old, new in rename_map.items():
        if old in df.columns:
            df = df.rename(columns={old: new})

    # parse de datas (se existirem)
    if "Dt_Entrega" in df.columns:
        df["Dt_Entrega"] = pd.to_datetime(df["Dt_Entrega"], dayfirst=True, errors="coerce")
    if "Dt_Ref" in df.columns:
        df["Dt_Ref"] = pd.to_datetime(df["Dt_Ref"], dayfirst=True, errors="coerce")

    if "Nome_Fundo" not in df.columns:
        # se não existir, tenta achar algo parecido
        col_guess = guess_column(df, ["nome do fundo", "fundo", "nome"])
        if col_guess:
            df = df.rename(columns={col_guess: "Nome_Fundo"})

    # filtro por nomes (exato após normalização)
    df["nome_norm"] = df.get("Nome_Fundo", "").apply(normalize)
    target_set = {normalize(x) for x in NOMES_RAW}
    filtered = df[df["nome_norm"].isin(target_set)].copy()

    # filtro por FRASES (categoria)
    pattern_cat = build_contains_pattern(FRASES_CATEGORIA)
    cat_col = None
    for key in ["categoria", "espécie", "especie", "assunto"]:
        cat_col = guess_column(filtered, [key])
        if cat_col:
            break

    if pattern_cat:
        if cat_col:
            filtered["cat_norm"] = filtered[cat_col].apply(normalize)
            result = filtered[filtered["cat_norm"].str.contains(pattern_cat, na=False)].copy()
        else:
            # fallback: busca em todas colunas texto
            text_cols = [c for c in filtered.columns if filtered[c].dtype == "object"]
            filtered["all_text_norm"] = filtered[text_cols].apply(
                lambda r: " ".join(normalize(v) for v in r.values if normalize(v)),
                axis=1
            )
            result = filtered[filtered["all_text_norm"].str.contains(pattern_cat, na=False)].copy()
    else:
        result = filtered.copy()

    # filtro por CHAVES (tipo)
    pattern_tipo = build_contains_pattern(CHAVES_TIPO)
    tipo_col = guess_column(result, ["tipo", "documento", "descricao", "descrição"])

    if pattern_tipo:
        if tipo_col:
            result["tipo_norm"] = result[tipo_col].apply(normalize)
            final = result[result["tipo_norm"].str.contains(pattern_tipo, na=False)].copy()
        else:
            text_cols = [c for c in result.columns if result[c].dtype == "object"]
            result["all_text_norm2"] = result[text_cols].apply(
                lambda r: " ".join(normalize(v) for v in r.values if normalize(v)),
                axis=1
            )
            final = result[result["all_text_norm2"].str.contains(pattern_tipo, na=False)].copy()
    else:
        final = result.copy()

    # remove colunas temporárias se existirem
    for tmp in ["nome_norm", "cat_norm", "tipo_norm", "all_text_norm", "all_text_norm2"]:
        if tmp in final.columns:
            final = final.drop(columns=[tmp])

    return final


def send_email(links: list[str]):
    user = os.getenv("GMAIL_USER")
    app_pass = os.getenv("GMAIL_APP_PASS")

    if not user or not app_pass:
        print("\n[AVISO] Variáveis de ambiente não configuradas. Não vou enviar email.")
        print("Configure assim (exemplo):")
        print("  export GMAIL_USER='seuemail@gmail.com'")
        print("  export GMAIL_APP_PASS='sua_app_password_de_16_chars'")
        print("\nLinks gerados:")
        print("\n".join(links))
        return

    text_body = "Olá,\n\nSeguem os links dos documentos:\n\n" + "\n".join(links)
    yag = yagmail.SMTP(user, app_pass)
    yag.send(to=EMAIL_TO, subject=EMAIL_SUBJECT, contents=[text_body])
    yag.close()
    print(f"Email enviado com {len(links)} links.")


def main():
    driver = make_driver()
    try:
        driver.get(URL)

        wait = WebDriverWait(driver, 20)

        # abrir filtros
        wait.until(EC.element_to_be_clickable((By.ID, "showFiltros"))).click()

        # dataInicial = ontem
        hoje = date.today()
        ontem = hoje - timedelta(days=1)
        valor_data = f"{ontem.day:02d}/{ontem.month:02d}/{ontem.year}"

        el = wait.until(EC.element_to_be_clickable((By.ID, "dataInicial")))
        el.clear()
        el.send_keys(valor_data)

        ok = select2_by_text_click(driver, "s2id_tipoFundo", "Fundo Imobiliário")
        if not ok:
            raise RuntimeError("Não consegui selecionar 'Fundo Imobiliário' no select2 (s2id_tipoFundo).")

        # filtrar
        wait.until(EC.element_to_be_clickable((By.ID, "filtrar"))).click()

        # esperar tabela e setar 100 linhas
        wait_table_ready(driver)
        dropdown = wait.until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "div#tblDocumentosEnviados_length select"))
        )
        Select(dropdown).select_by_value("100")
        wait_table_ready(driver)

        # coletar páginas
        data = collect_pages(driver)
        if data.empty:
            print("Tabela vazia / nada encontrado.")
            return

    finally:
        try:
            driver.quit()
        except Exception:
            pass

    # filtrar e preparar links
    df_final = filter_df(data)

    if df_final.empty:
        print("Nenhum registro após os filtros.")
        return

    # montar links
    links = []
    for row in df_final.itertuples(index=False):
        doc = getattr(row, "DocNumber", None)
        if doc is None or (isinstance(doc, float) and pd.isna(doc)) or str(doc).strip() == "":
            continue

        # tenta pegar Categoria/Tipo se existirem
        cat = getattr(row, "Categoria", None) if "Categoria" in df_final.columns else None
        tipo = getattr(row, "Tipo", None) if "Tipo" in df_final.columns else None

        parts = []
        if cat is not None and not (isinstance(cat, float) and pd.isna(cat)) and str(cat).strip():
            parts.append(str(cat).strip())
        if tipo is not None and not (isinstance(tipo, float) and pd.isna(tipo)) and str(tipo).strip():
            parts.append(str(tipo).strip())

        prefix = " - ".join(parts) if parts else "Sem categoria/tipo"
        link = f"https://fnet.bmfbovespa.com.br/fnet/publico/visualizarDocumento?id={doc}&cvm=true"
        links.append(f"{prefix}: {link}")

    # salva CSVs
    df_final.to_csv("resultado_filtrado.csv", index=False)
    print(f"Salvo: resultado_filtrado.csv ({len(df_final)} linhas)")

    # envia email (ou imprime links se não tiver env vars)
    send_email(links)


if __name__ == "__main__":
    main()
