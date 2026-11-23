import pandas as pd
from openpyxl import load_workbook
from ofxparse import OfxParser
import os
import glob
from datetime import date, timedelta
import io
import unicodedata
import numpy as np
import streamlit as st
import io


data_hj = date.today()
data_ontem = int((data_hj - timedelta(days=3)).strftime("%d"))
revendas_contas = {
    "BOA VISTA": "2035-4",
    "ITAPIPOCA":"1996-8",
    "CRUZ":"2000-1"
}
transacoes = []
arquivinhos = st.file_uploader(label="Coloque aqui os arquivos OFX, pai", accept_multiple_files=True, type=["OFX"])
if not arquivinhos:
 st.stop()
def conci():
    for chave, conta in revendas_contas.items():
     for extrato in arquivinhos:
          ler_acento = p.read().decode('cp1252')
          sem_acento = unicodedata.normalize("NFKD", ler_acento).encode("ascii", "ignore").decode("ascii")
          ofx = OfxParser.parse(io.StringIO(sem_acento))
          try:
           for acoount in ofx.accounts:
             stmt = acoount.statement
             saldo = getattr(stmt, 'ledger_balance', None)
             if saldo is None:
                  saldo = getattr(stmt, 'balance', 0.0)
             for t in acoount.statement.transactions:
               transacoes.append({
                   "Data": t.date,
                   "Beneficario/descrição": t.memo,
                   "Valor": t.amount,
                   "Conta": conta,
                   "Saldo": saldo
               })
         except Exception as e:
            print("deu tudo errado! Corre!")
    
conci()
df_ = pd.DataFrame(transacoes)


gastos = {"PIX RECEBIDO": "PIX", "LIQUIDACAO DE COBRANCA VALOR DISPONIVEL": "BOLETO", "RENTAB.INVEST FACILCRED": "Rendimentos", "PIX QR CODE DINAMICO DES: AMBEV S A": "AMBEV NOTAS", ("PGTO SALARIO VIA NET EMP", "PGTO FERIAS VIA NET EMPR", "PGTO GRATIFICACAO/PREMIO"): "FOLHA DE PAGAMENTO"}
def extrato(lancamento):
  textin = str(lancamento)
  for chave, valor in gastos.items():
    if isinstance(chave, tuple):
      nomes = chave
    else:
      nomes = (chave,)
    for nome in nomes:
      if nome in textin:
        return valor
  return "Outros"

df_conci = df_
separar_revenda = {
  "REVENDA AM": "2035-4",
  "CAPITALI ITAPIPOCA": "1996-8",
  "CAPITALI CRUZ":"2000-1"
}

condicao = []
revenda_in = []
for nome, conta in separar_revenda.items():
  condicao.append(df_conci["Conta"] == conta)
  revenda_in.append(nome)
colunas = ["Valor", "Saldo"]
for col in df_conci[colunas]:
  df_conci[col] = pd.to_numeric(df_conci[col], errors="coerce")
df_conci["Data"] = pd.to_datetime(df_conci["Data"]).dt.strftime("%d/%m/%Y")
df_conci["Tipo de Entrada/Saida"] = df_conci["Beneficario/descrição"].apply(extrato)
df_conci["Revenda"] = np.select(condicao, revenda_in, default="Outros")
df_conci = df_conci.drop(columns=["Conta"])
df_conci = df_conci.drop(columns=["Saldo"])
with pd.ExcelWriter(banco, engine="openpyxl", mode="w") as w:
  abas_rev = ["CAPITALI CRUZ", "CAPITALI ITAPIPOCA", "REVENDA AM"]
  for aba in abas_rev:
    df_filtrado = df_conci[df_conci["Revenda"] == aba]
    df_filtrado.to_excel(w, sheet_name=aba, index=False)


st.set_page_config(layout="wide")
st.title("Conciliação - Banco:")
revendas__ = ["CAPITALI CRUZ", "CAPITALI ITAPIPOCA", "REVENDA AM"]
revenda_fil = st.sidebar.selectbox("Selecione a Revenda:", revendas__)
df_final = pd.read_excel(banco, sheet_name=revenda_fil)
datas_conci = sorted(df_final["Data"].dropna().unique())
try:
 if datas_conci:
   data_escolhida = st.sidebar.selectbox("Selecione as datas:", datas_conci)
   df_filtrado = df_final[(df_final["Data"] == data_escolhida) & (df_final["Tipo de Entrada/Saida"].isin(["PIX", "BOLETO", "Rendimentos", "AMBEV NOTAS", "FOLHA DE PAGAMENTO"]))]
   filtropix_boleto = df_filtrado.groupby("Tipo de Entrada/Saida")["Valor"].sum()
   boleto_ = filtropix_boleto.get("BOLETO", 0.0)
   pix_ = filtropix_boleto.get("PIX", 0.0)
   rend_ = filtropix_boleto.get("Rendimentos", 0.0)
   ambev_ = filtropix_boleto.get("AMBEV NOTAS", 0.0)
   folhas_ = filtropix_boleto.get("FOLHA DE PAGAMENTO", 0.0)



   metricas_geral = st.columns(5)
   metricas_geral[0].metric(label="Receita - Boleto:", value=f"{boleto_: ,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
   metricas_geral[1].metric(label="Receita - PIX:", value=f"{pix_: ,.2f}".replace(",", "X").replace(".", ",").replace("X", "."))
   metricas_geral[2].metric(label="Rendimento e Aplicação:", value=f"{rend_: ,.3f}".replace(",", "X").replace(".", ",").replace("X", "."))
   metricas_geral[3].metric(label="Valor das notas ambev:", value=f"{ambev_: ,.3f}".replace(",", "X").replace(".", ",").replace("X", "."))
   metricas_geral[4].metric(label="Folha de pagamento:", value=f"{folhas_: ,.3f}".replace(",", "X").replace(".", ",").replace("X", "."))


   st.dataframe(df_filtrado, hide_index=True, use_container_width=True)
 else:
   print("k!")
except Exception as e:
  print("deu tudo errado nessa bosta!")


st.divider()
st.title(f"Entradas - {data_escolhida}")
df_filtrado_entrada = df_final[(df_final["Valor"] > 0) & (df_final["Data"] == data_escolhida)]
st.dataframe(df_filtrado_entrada, hide_index=True, use_container_width=True)
apenas_entradas = io.BytesIO()
with pd.ExcelWriter(apenas_entradas, engine="openpyxl") as en:
  abas_rev = ["CAPITALI CRUZ", "CAPITALI ITAPIPOCA", "REVENDA AM"]
  for aba in abas_rev:
   df_entradas = df_filtrado_entrada[df_filtrado_entrada["Revenda"] == aba]
   df_entradas.to_excel(en, sheet_name=aba, index=False)
st.download_button(label="baixe aqui apenas as entradas do periodo:", data=apenas_entradas.getvalue(), file_name=f"entradas de {revenda_fil} - {data_escolhida}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

st.divider()
st.title(f"Saidas - {data_escolhida}")
df_filtrado_saida = df_final[(df_final["Valor"] < 0) & (df_final["Data"] == data_escolhida)]
st.dataframe(df_filtrado_saida, hide_index=True, use_container_width=True)
apenas_saidas = io.BytesIO()
with pd.ExcelWriter(apenas_saidas, engine="openpyxl") as s:
  abas_rev = ["CAPITALI CRUZ", "CAPITALI ITAPIPOCA", "REVENDA AM"]
  for aba in abas_rev:
   df_saidas = df_filtrado_saida[df_filtrado_saida["Revenda"] == aba]
   df_saidas.to_excel(s, sheet_name=aba, index=False)
st.download_button(label="baixe aqui apenas as saidas do periodo para uma melhor visualização:", data=apenas_saidas.getvalue(), file_name=f"saidas de {revenda_fil} - {data_escolhida}.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")


