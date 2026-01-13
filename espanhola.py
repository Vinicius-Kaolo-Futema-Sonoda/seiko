import os
import yfinance as yf
import pandas as pd
from datetime import datetime
import yagmail

# === CONFIG DO EMAIL ===
EMAIL_TO = ["diego@seikopartners.com.br", "fabio@seikopartners.com.br"]
EMAIL_SUBJECT = "Relatório dos Fundo Imobiliários Espanhola"

# 1) Tickers
tickers = ['AMS.MC', 'APAM.MC', 'CLNX.MC', 'ELE.MC', 'IUSC.DE', "SAJA.F"]
fx_ticker = 'EURBRL=X'

# 2) Baixar dados desde 30/09/2025
data = yf.download(
    tickers + [fx_ticker],
    start="2025-09-30"
)

# 3) Pegar apenas o Close
close = data["Close"]

# separar ações e câmbio
close_ativos = close[tickers]           # colunas das ações
eurbrl = close[fx_ticker]               # série do câmbio

# 4) Converter para BRL (preço em EUR * EURBRL)
close_brl = close_ativos.mul(eurbrl, axis=0)

# 5) Ajustar índice para coluna e renomear
close_brl = close_brl.reset_index()     # índice Date vira coluna
close_brl.rename(columns={"Date": "Data"}, inplace=True)

# >>> Remover a hora da data <<<
close_brl["Data"] = close_brl["Data"].dt.date

# 6) Usar melt para formato longo
df_melt = close_brl.melt(
    id_vars="Data",
    var_name="Série",
    value_name="Valor"
)

def exportar_excel(close_brl: pd.DataFrame, df_melt: pd.DataFrame, out_path: str) -> str:
    """Exporta um XLSX com duas abas e retorna o caminho."""
    with pd.ExcelWriter(out_path, engine="openpyxl") as writer:
        close_brl.to_excel(writer, index=False, sheet_name="precos_brl")
        df_melt.to_excel(writer, index=False, sheet_name="melt")
    return out_path

def send_email(lines: list[str], attachments: list[str] | None = None):
    user = "viniciuskaolo@gmail.com"           # seu gmail
    app_pass = "hjlvknog yxjtmuku"   # app password (16 chars)

    if not user or not app_pass:
        print("\n[AVISO] Variáveis de ambiente não configuradas. Não vou enviar email.")
        print("Configure assim (exemplo):")
        print("  export GMAIL_USER='seuemail@gmail.com'")
        print("  export GMAIL_APP_PASS='sua_app_password_de_16_chars'")
        print("\nConteúdo que iria no email:\n")
        print("\n".join(lines))
        if attachments:
            print("\nAnexos que eu mandaria:")
            for a in attachments:
                print(" -", a)
        return

    text_body = "Olá,\n\nSeguem os documentos:\n\n" + "\n".join(lines)

    yag = yagmail.SMTP(user, app_pass)
    yag.send(
        to=EMAIL_TO,
        subject=EMAIL_SUBJECT,
        contents=[text_body],
        attachments=attachments or []
    )
    yag.close()
    print(f"Email enviado com {len(lines)} linhas e {len(attachments or [])} anexo(s).")

# === GERA O EXCEL E ENVIA ===
timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
xlsx_path = f"cotacoes_brl_{timestamp}.xlsx"

exportar_excel(close_brl, df_melt, xlsx_path)

linhas_email = [
    f"Arquivo Excel: {os.path.abspath(xlsx_path)}",
    f"Tickers: {', '.join(tickers)}",
    f"Câmbio: {fx_ticker}",
    "Abas: precos_brl (largo) e melt (longo)"
]

send_email(linhas_email, attachments=[xlsx_path])
