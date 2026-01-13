from __future__ import annotations

import os
import re
import unicodedata
from datetime import date, timedelta

import pandas as pd
import yagmail
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import Select, WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from selenium.common.exceptions import StaleElementReferenceException, TimeoutException, ElementClickInterceptedException


URL = "https://fnet.bmfbovespa.com.br/fnet/publico/abrirGerenciadorDocumentosCVM"

# (nome do fundo, código/ticker)
NOMES_RAW: list[tuple[str, str]] = [
    ("CARTESIA RECEBÍVEIS IMOBILIÁRIOS FII", "CACR11"),
    ("FII AF INVEST CRI", "AFHI11"),
    ("ALIANZA TRUST RI FII", "ALZR11"),
    ("RIZA ARCTIUM FII", "RZAT11"),
    ("BRC RENDA CORPORATIVA FII RESP LIMITADA", "FATN11"),
    ("AZ QUEST PANORAMA LOGISTICA", "AZPL11"),
    ("BB PREMIUM MALLS FII DE RESPONSABILIDADE LIMITADA", "BBIG11"),
    ("BC FUND FII", "BRCR11"),
    ("FII GREEN TOWERS RESPONSABILIDADE LIMITADA", "BCIA11"),
    ("BRADESCO CARTEIRA IMOBILIARIA ATIVA FII", "BCIA11"),
    ("BANESTES RECEBIVEIS IMOBILIARIO FII", "BCRI11"),
    ("FII BLUEMA", "BLMG11"),
    ("BRESCO LOGISTICA FUNDO DE INVESTIMENTO IMOBILIÁRIO", "BRCO11"),
    ("BRPR CORPORATE FII", "BROF11"),
    ("BTGP AGRO LOG FII", "BTAL11"),
    ("FUNDO DE CRI  FII", "BTCI11"),
    ("FII BTG PACTUAL SHOPPINGS", "BPML11"),
    ("BTGP HEDGE FUND FII", "BTHF11"),
    ("CANUMA CAPITAL FII", "CCME11"),
    ("ITAÚ CRÉDITO IMOBILIÁRIO IPCA FII", "ICRI11"),
    ("CLAVE INDICES FII", "CLIN11"),
    ("FII SHOPPINGS AAA", "CPSH11"),
    ("CYRELA CRÉDITO - FUNDO DE INVESTIMENTO IMOBILIÁRIO", "CYCR11"),
    ("FII DEVANT", "DEVA11"),
    ("FATOR VERITÀ FUNDO DE INVESTIMENTO IMOBILIÁRIO FII", "VRTA11"),
    ("GAZIT MALLS FII", "GZIT11"),
    ("GGR COVEPI RENDA - FII - RESPONSABILIDADE LIMITADA", "GGRC11"),
    ("ICATU VANG GRU LOG FII", "GRUL11"),
    ("FII HABITA", "HABT11"),
    ("FII HECTAR", "HCTR11"),
    ("HEDGE BRASIL SHOPPING FUNDO DE INVESTIMENTO IMOBILIÁRIO", "HGBS11"),
    ("PÁTRIA RECEBÍVEIS IMOBILIÁRIOS - FII", "HGCR11"),
    ("PÁTRIA LOG - FUNDO DE INVESTIMENTO IMOBILIÁRIO", "HGLG11"),
    ("PÁTRIA ESCRITÓRIOS FUNDO DE INVESTIMENTO IMOBILIÁRIO", "HGRE11"),
    ("PÁTRIA RENDA URBANA - FII", "HGRU11"),
    ("HOTEL MAXINVEST FII", "HTMX11"),
    ("HSI Ativos Financeiros FII", "HSAF11"),
    ("HSI LOGÍSTICA FII", "HSLG11"),
    ("HSI MALLS FUNDO DE INVESTIMENTO IMOBILIÁRIO", "HSML11"),
    ("HEDGE TOP FOFII 3 FUNDO DE INVESTIMENTO IMOBILIÁRIO", "HFOF11"),
    ("ITAÚ TOTAL RETURN FII", "ITRI11"),
    ("JS ATIVOS FINANCEIROS FUNDO DE INVESTIMENTO IMOBILIÁRIO", "JSAF11"),
    ("JS REAL ESTATE MULTIGESTÃO", "JSRE11"),
    ("JS RECEBÍVEIS IMOBILIÁRIOS FI IMOBILIÁRIO", "JSCR11"),
    ("KILIMA SUNO 30 FII", "KISU11"),
    ("KINEA RENDA IMOBILIÁRIA FII", "KNRI11"),
    ("KINEA CREDITAS FUNDO DE INVESTIMENTO IMOBILIÁRIO – FII", "KCRE11"),
    ("KINEA HEDGE FUND FDO INV IMOBILIÁRIO", "KNHF11"),
    ("KINEA HIGH YIELD CRI FII", "KNHY11"),
    ("KINEA ÍNDICE DE PREÇOS FII", "KNIP11"),
    ("KINEA SECURITIES FII", "KNSC11"),
    ("KINEA UNIQUE HY CDI FDO INV IMOBILIÁRIO", "KNUQ11"),
    ("FUNDO DE FUNDOS DE INVESTIMENTO IMOBILIÁRIO KINEA FII", "KFOF11"),
    ("KILIMA VOLKANO RECEBIVEIS IMOBILIARIOS FII", "KIVO11"),
    ("KINEA OPORTUNIDADES REAL ESTATE FII", "KORE11"),
    ("LIFE CAPITAL PARTNERS FUNDO DE INVESTIMENTOS IMOBILIÁRIOS", "LIFE11"),
    ("VBI LOGISTICO FII", "LVBI11"),
    ("FII MANATÍ CAPITAL", "MANA11"),
    ("MAUA CAPITAL RE FII", "MCCI11"),
    ("MAUÁ HIGH YIELD FII", "MCRE11"),
    ("MAXI RENDA FII", "MXRF11"),
    ("MERITO DESENVOLVIMENTO IMOBILIARIO I FII", "MFII11"),
    ("OURINVEST JPP FUNDO DE INVESTIMENTO IMOBILIÁRIO – FII", "OUJP11"),
    ("PATRIA LOG", "PATL11"),
    ("PATRIA CRÉDITO IMOBILIÁRIO ÍNDICE DE PREÇOS FII", "PCIP11"),
    ("PARAMIS HEDGE FUND FII", "PMIS11"),
    ("PATRIA MALLS FUNDO DE INVESTIMENTO IMOBILIÁRIO", "PMLL11"),
    ("POLO CRÉDITOS IMOBILIÁRIOS FII", "PORD11"),
    ("PATRIA SECURITIES FII", "PSEC11"),
    ("FII VBI PRIME PROPERTIES", "PVBI11"),
    ("RBR LOG - FUNDO DE INVESTIMENTO IMOBILIARIO", "RBRL11"),
    ("RBR PLUS MULTI FII", "RBRX11"),
    ("FII RBR HIGH YIELD", "RBRY11"),
    ("RBR PREMIUM RECEBÍVEIS IMOBILIÁRIOS FII", "RPRI11"),
    ("RBR PROPERTIES - FII", "RBRP11"),
    ("RBR HIGH GRADE FII", "RBRR11"),
    ("FII REC RECEBÍVEIS IMOBILIÁRIOS", "RECR11"),
    ("FII REC RENDA IMOBILIÁRIA", "RECT11"),
    ("RIO BRAVO FUNDO DE FUNDOS DE INVESTIMENTO IMOBILIÁRIO", "RBFF11"),
    ("FII RIO BRAVO RENDA CORPORATIVA", "RCRB11"),
    ("FII RIO BRAVO RENDA VAREJO - FII", "RBVA11"),
    ("RIZA AKIN FII", "RZAK11"),
    ("Fundo de Investimento Imobiliário Riza Terrax", "RZTR11"),
    ("Tellus Rio Bravo Renda Logística FII Resp. Ltda", "TRBL11"),
    ("SPX SYN MULTI FII", "SPXS11"),
    ("SUNO CRI FII", "SNCI11"),
    ("SUNO FOF FII", "SNFF11"),
    ("FII TELLUS PROPERTIES", "TEPP11"),
    ("TG ATIVO R", "TGAR11"),
    ("TIVIO RENDA IMOBILIÁRIA FII", "TVRI11"),
    ("FII RBR TOP OFFICES", "TOPP11"),
    ("FII URCA", "URPR11"),
    ("VALORA HEDGE FUND FII", "VGHF11"),
    ("VALORA CRI ÍNDICE FII", "VGIP11"),
    ("VALORA RE III FII", "VGIR11"),
    ("VECTIS JUROS REAL FII", "VCJR11"),
    ("VALORA RENDA IMOBILIARIA FII", "VGRI11"),
    ("Vinci Imóveis Urbanos Fundo de Investimento Imobiliário", "VIUR11"),
    ("VINCI LOGISTICA FII", "VILG11"),
    ("VINCI OFFICES FII", "VINO11"),
    ("VINCI SHOPPING CENTERS FII", "VISC11"),
    ("FATOR VERITÀ MULTIESTRATEGIA FII", "VRTM11"),
    ("WHG REAL ESTATE FUNDO DE INVESTIMENTO IMOBILIÁRIO", "WHGR11"),
    ("XP CRÉDITO IMOBILIÁRIO - FUNDO DE INVESTIMENTO IMOBILIÁRIO", "XPCI11"),
    ("XP MALLS FUNDO DE INVESTIMENTO IMOBILIÁRIOS FII", "XPML11"),
    ("XP SELECTION FUNDO DE FUNDOS  - FII", "XPSF11"),
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
    s = re.sub(r"[\u0300-\u036f]", "", s)
    s = re.sub(r"\s+", " ", s)
    return s


def build_contains_pattern(terms: list[str]) -> str:
    terms_norm = [normalize(t) for t in terms if t and str(t).strip()]
    terms_norm = [t for t in terms_norm if t]
    if not terms_norm:
        return ""
    return "|".join(re.escape(t) for t in terms_norm)


def guess_column(df: pd.DataFrame, keywords: list[str]) -> str | None:
    for col in list(df.columns):
        col_norm = normalize(col)
        if any(normalize(k) in col_norm for k in keywords):
            return col
    return None


def select2_by_text_click(driver, container_id: str, option_text: str, timeout: int = 15) -> bool:
    container = WebDriverWait(driver, timeout).until(
        EC.element_to_be_clickable((By.ID, container_id))
    )
    container.click()

    # select2 v3
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

    # select2 v4
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
    opts.add_argument("--headless=new")
    opts.add_argument("--no-sandbox")
    opts.add_argument("--disable-dev-shm-usage")
    opts.add_argument("--window-size=1920,1080")
    return webdriver.Chrome(options=opts)


def wait_table_ready(driver, timeout: int = 30):
    WebDriverWait(driver, timeout).until(
        EC.presence_of_element_located((By.ID, "tblDocumentosEnviados"))
    )
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

    if "Ações" in df.columns:
        df["DocNumber"] = (
            df["Ações"]
            .astype(str)
            .str.replace("visualizarDocumento?id=", "", regex=False)
            .str.replace("&cvm=true", "", regex=False)
        )
        df = df.drop(columns=["Ações"])

    return df


def _safe_click(driver, element):
    try:
        element.click()
        return
    except (ElementClickInterceptedException, StaleElementReferenceException):
        pass
    # fallback: scroll + JS click
    driver.execute_script("arguments[0].scrollIntoView({block:'center'});", element)
    driver.execute_script("arguments[0].click();", element)


def collect_pages(driver) -> pd.DataFrame:
    """
    Paginação robusta (evita stale):
    - lê a página atual
    - enquanto "next" não estiver disabled:
        - guarda o primeiro <tr> atual (para esperar staleness)
        - re-encontra o botão next a cada iteração
        - clica com segurança
        - espera o <tr> antigo ficar stale (tabela trocou)
        - espera tabela pronta
        - lê
    """
    pages: list[pd.DataFrame] = []

    wait_table_ready(driver)
    pages.append(read_table(driver))

    while True:
        # re-encontra o botão next SEMPRE (evita stale)
        try:
            next_btn = driver.find_element(By.ID, "tblDocumentosEnviados_next")
            classes = next_btn.get_attribute("class") or ""
            if "disabled" in classes:
                break
        except StaleElementReferenceException:
            # se deu stale aqui, tenta de novo
            continue

        # pega o primeiro row atual para esperar a troca da tabela
        try:
            old_first_row = driver.find_element(By.CSS_SELECTOR, "#tblDocumentosEnviados tbody tr")
        except Exception:
            old_first_row = None

        # tenta clicar (com retry leve)
        clicked = False
        for _ in range(3):
            try:
                next_btn = driver.find_element(By.ID, "tblDocumentosEnviados_next")
                _safe_click(driver, next_btn)
                clicked = True
                break
            except StaleElementReferenceException:
                continue

        if not clicked:
            # não conseguiu avançar, para evitar loop infinito
            break

        # espera a tabela realmente trocar
        if old_first_row is not None:
            try:
                WebDriverWait(driver, 15).until(EC.staleness_of(old_first_row))
            except TimeoutException:
                # se não ficou stale, ainda assim tentamos esperar a tabela
                pass

        wait_table_ready(driver)
        pages.append(read_table(driver))

    if not pages:
        return pd.DataFrame()

    return pd.concat(pages, ignore_index=True)


def filter_df(df: pd.DataFrame) -> pd.DataFrame:
    rename_map = {
        "Nome do Fundo": "Nome_Fundo",
        "Data de Referência": "Dt_Ref",
        "Data de Entrega": "Dt_Entrega",
        "Espécie": "Especie",
    }
    for old, new in rename_map.items():
        if old in df.columns:
            df = df.rename(columns={old: new})

    if "Dt_Entrega" in df.columns:
        df["Dt_Entrega"] = pd.to_datetime(df["Dt_Entrega"], dayfirst=True, errors="coerce")
    if "Dt_Ref" in df.columns:
        df["Dt_Ref"] = pd.to_datetime(df["Dt_Ref"], dayfirst=True, errors="coerce")

    if "Nome_Fundo" not in df.columns:
        col_guess = guess_column(df, ["nome do fundo", "fundo", "nome"])
        if col_guess:
            df = df.rename(columns={col_guess: "Nome_Fundo"})

    nome_para_codigo = {normalize(nome): codigo for (nome, codigo) in NOMES_RAW}

    df["nome_norm"] = df.get("Nome_Fundo", "").apply(normalize)
    filtered = df[df["nome_norm"].isin(nome_para_codigo.keys())].copy()
    filtered["Codigo_Fundo"] = filtered["nome_norm"].map(nome_para_codigo).fillna("")

    # tenta padronizar "Categoria" e "Tipo"
    cat_guess = guess_column(filtered, ["categoria", "espécie", "especie", "assunto"])
    if cat_guess and cat_guess != "Categoria":
        filtered = filtered.rename(columns={cat_guess: "Categoria"})
    tipo_guess = guess_column(filtered, ["tipo", "documento", "descricao", "descrição"])
    if tipo_guess and tipo_guess != "Tipo":
        filtered = filtered.rename(columns={tipo_guess: "Tipo"})

    # filtro por categoria
    pattern_cat = build_contains_pattern(FRASES_CATEGORIA)
    if pattern_cat:
        if "Categoria" in filtered.columns:
            filtered["cat_norm"] = filtered["Categoria"].apply(normalize)
            filtered = filtered[filtered["cat_norm"].str.contains(pattern_cat, na=False)].copy()
        else:
            text_cols = [c for c in filtered.columns if filtered[c].dtype == "object"]
            filtered["all_text_norm"] = filtered[text_cols].apply(
                lambda r: " ".join(normalize(v) for v in r.values if normalize(v)),
                axis=1
            )
            filtered = filtered[filtered["all_text_norm"].str.contains(pattern_cat, na=False)].copy()

    # filtro por tipo
    pattern_tipo = build_contains_pattern(CHAVES_TIPO)
    if pattern_tipo:
        if "Tipo" in filtered.columns:
            filtered["tipo_norm"] = filtered["Tipo"].apply(normalize)
            filtered = filtered[filtered["tipo_norm"].str.contains(pattern_tipo, na=False)].copy()
        else:
            text_cols = [c for c in filtered.columns if filtered[c].dtype == "object"]
            filtered["all_text_norm2"] = filtered[text_cols].apply(
                lambda r: " ".join(normalize(v) for v in r.values if normalize(v)),
                axis=1
            )
            filtered = filtered[filtered["all_text_norm2"].str.contains(pattern_tipo, na=False)].copy()

    for tmp in ["nome_norm", "cat_norm", "tipo_norm", "all_text_norm", "all_text_norm2"]:
        if tmp in filtered.columns:
            filtered = filtered.drop(columns=[tmp])

    return filtered


def send_email(lines: list[str]):
    user = "viniciuskaolo@gmail.com"           # seu gmail
    app_pass = "hjlv knog yxjt muku"   # app password (16 chars)

    if not user or not app_pass:
        print("\n[AVISO] Variáveis de ambiente não configuradas. Não vou enviar email.")
        print("Configure assim (exemplo):")
        print("  export GMAIL_USER='seuemail@gmail.com'")
        print("  export GMAIL_APP_PASS='sua_app_password_de_16_chars'")
        print("\nConteúdo que iria no email:\n")
        print("\n".join(lines))
        return

    text_body = "Olá,\n\nSeguem os documentos:\n\n" + "\n".join(lines)
    yag = yagmail.SMTP(user, app_pass)
    yag.send(to=EMAIL_TO, subject=EMAIL_SUBJECT, contents=[text_body])
    yag.close()
    print(f"Email enviado com {len(lines)} linhas.")


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
            raise RuntimeError("Não consegui selecionar 'Fundo Imobiliário'.")

        # filtrar
        wait.until(EC.element_to_be_clickable((By.ID, "filtrar"))).click()

        # esperar tabela e setar 100 linhas
        wait_table_ready(driver)
        dropdown = wait.until(
            EC.presence_of_element_located((By.CSS_SELECTOR, "div#tblDocumentosEnviados_length select"))
        )
        Select(dropdown).select_by_value("100")
        wait_table_ready(driver)

        data = collect_pages(driver)
        if data.empty:
            print("Tabela vazia / nada encontrado.")
            return

    finally:
        try:
            driver.quit()
        except Exception:
            pass

    df_final = filter_df(data)

    if df_final.empty:
        print("Nenhum registro após os filtros.")
        return

    # linhas: CODIGO | CATEGORIA - TIPO: LINK
    lines: list[str] = []
    for row in df_final.itertuples(index=False):
        doc = getattr(row, "DocNumber", None)
        if doc is None or (isinstance(doc, float) and pd.isna(doc)) or str(doc).strip() == "":
            continue

        codigo = getattr(row, "Codigo_Fundo", "") or ""
        if isinstance(codigo, float) and pd.isna(codigo):
            codigo = ""
        codigo = str(codigo).strip()

        categoria = getattr(row, "Categoria", "") if "Categoria" in df_final.columns else ""
        if isinstance(categoria, float) and pd.isna(categoria):
            categoria = ""
        categoria = str(categoria).strip()

        tipo = getattr(row, "Tipo", "") if "Tipo" in df_final.columns else ""
        if isinstance(tipo, float) and pd.isna(tipo):
            tipo = ""
        tipo = str(tipo).strip()

        prefix = " - ".join([p for p in [categoria, tipo] if p]) or "Sem categoria/tipo"
        link = f"https://fnet.bmfbovespa.com.br/fnet/publico/visualizarDocumento?id={doc}&cvm=true"

        lines.append(f"{codigo} | {prefix}: {link}")

    df_final.to_csv("resultado_filtrado.csv", index=False)
    print(f"Salvo: resultado_filtrado.csv ({len(df_final)} linhas)")

    send_email(lines)


if __name__ == "__main__":
    main()
