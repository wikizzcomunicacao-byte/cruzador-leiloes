import io
import re
import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="Cruzamento de Leilões", page_icon="🏢", layout="wide"
)

st.title("🏢 Cruzador Automático de Leilões x Investidores")
st.markdown(
    "Faça o upload das duas planilhas em Excel para gerar o cruzamento de"
    " oportunidades."
)


def normalize(text):
  if not isinstance(text, str):
    return ""
  text = text.lower()
  text = re.sub(r"[áàâãä]", "a", text)
  text = re.sub(r"[éèêë]", "e", text)
  text = re.sub(r"[íìîï]", "i", text)
  text = re.sub(r"[óòôõö]", "o", text)
  text = re.sub(r"[úùûü]", "u", text)
  text = re.sub(r"[ç]", "c", text)
  text = re.sub(r"[^a-z0-9\s]", " ", text)
  return " ".join(text.split())


def parse_budget(budget_str):
  budget_str = str(budget_str)
  if "100.000" in budget_str:
    return 0, 100000
  elif "200.000" in budget_str:
    return 0, 200000
  elif "300.000" in budget_str:
    return 0, 300000
  elif "400.000" in budget_str:
    return 0, 400000
  elif "500.000" in budget_str:
    return 500000, 999999999
  else:
    return 0, 999999999


def parse_types(tipos_str):
  norm_t = normalize(tipos_str)
  res = set()
  if "apartamento" in norm_t:
    res.add("Apartamento")
  if "casa" in norm_t:
    res.add("Casa")
  if "terreno" in norm_t:
    res.add("Terreno")
  if "lote comercial" in norm_t or "comercial" in norm_t:
    res.add("Comercial")
  if "outros" in norm_t:
    res.update(["Area Rural", "Comercial", "Indefinido", "Vaga de Garagem"])
  return list(res)


col1, col2 = st.columns(2)

with col1:
  file_leiloes = st.file_uploader(
      "1️⃣ Planilha de Leilões (.xlsx)", type=["xlsx", "xls"]
  )

with col2:
  file_investidores = st.file_uploader(
      "2️⃣ Planilha de Investidores (.xlsx)", type=["xlsx", "xls"]
  )

if file_leiloes and file_investidores:
  if st.button("🚀 Executar Cruzamento de Dados", type="primary"):
    with st.spinner("Cruzando dados..."):
      try:
        df_leiloes = pd.read_excel(file_leiloes)
        df_investidores = pd.read_excel(file_investidores)

        df_leiloes["norm_cidade"] = df_leiloes["Cidade"].apply(normalize)
        df_leiloes["norm_estado"] = df_leiloes["Estado"].apply(normalize)
        df_leiloes["norm_tipo"] = df_leiloes["Tipo de Bem"].apply(normalize)

        df_leiloes["preco_effective"] = df_leiloes.apply(
            lambda r: (
                r["2º Leilão (Preço)"]
                if pd.notnull(r["2º Leilão (Preço)"])
                and r["2º Leilão (Preço)"] > 0
                else r["1º Leilão (Preço)"]
            ),
            axis=1,
        )

        df_leiloes["desconto_%"] = np.where(
            (pd.notnull(df_leiloes["Valor de Avaliação do Leiloeiro"]))
            & (df_leiloes["Valor de Avaliação do Leiloeiro"] > 0),
            ((
                df_leiloes["Valor de Avaliação do Leiloeiro"]
                - df_leiloes["preco_effective"]
            )
            / df_leiloes["Valor de Avaliação do Leiloeiro"])
            * 100,
            0,
        )

        cols = list(df_investidores.columns)
        resultados = []

        for idx, row in df_investidores.iterrows():
          nome = str(row["Nome Completo"]).strip()
          cidades_input = str(row[cols[10]]).strip()
          tipos_input = str(row[cols[7]]).strip()
          valor_input = str(row[cols[8]]).strip()
          cons = str(row[cols[15]]).strip()

          norm_cid = normalize(cidades_input)
          norm_cons = normalize(cons)
          norm_val = normalize(valor_input)

          sub = df_leiloes.copy()

          target_cidades = []
          target_estados = []

          if any(
              w in norm_cid or w in norm_cons
              for w in [
                  "rio preto",
                  "sao jose do rio preto",
                  "sao jose do preto",
                  "sjrp",
              ]
          ):
            target_cidades.append("sao jose do rio preto")
          if "bady" in norm_cid or "bady" in norm_cons:
            target_cidades.append("bady bassitt")
          if "mirassol" in norm_cid or "mirassol" in norm_cons:
            target_cidades.append("mirassol")
          if "santo andre" in norm_cid or "santo andre" in norm_cons:
            target_cidades.append("santo andre")
          if "sao bernardo" in norm_cid or "sao bernardo" in norm_cons:
            target_cidades.append("sao bernardo do campo")
          if "sao caetano" in norm_cid or "sao caetano" in norm_cons:
            target_cidades.append("sao caetano do sul")
          if "abc" in norm_cid or "abc" in norm_cons:
            target_cidades.extend(
                ["santo andre", "sao bernardo do campo", "sao caetano do sul"]
            )
          if "sao paulo" in norm_cid or "sao paulo" in norm_cons:
            if (
                "grande sao paulo" in norm_cid
                or "grande sao paulo" in norm_cons
            ):
              target_cidades.extend([
                  "sao paulo",
                  "santo andre",
                  "sao bernardo do campo",
                  "sao caetano do sul",
              ])
            elif "estado de sp" in norm_cid or "estado de sp" in norm_cons:
              target_estados.append("sao paulo")
            else:
              target_cidades.append("sao paulo")
          if "sorocaba" in norm_cid or "sorocaba" in norm_cons:
            target_cidades.append("sorocaba")
          if "tatui" in norm_cid or "tatui" in norm_cons:
            target_cidades.append("tatui")
          if "boituva" in norm_cid or "boituva" in norm_cons:
            target_cidades.append("boituva")
          if "rio de janeiro" in norm_cid or "rio de janeiro" in norm_cons:
            target_cidades.append("rio de janeiro")
          if "goiania" in norm_cid:
            target_cidades.append("goiania")
          if "brasilia" in norm_cid:
            target_cidades.append("brasilia")
          if "anapolis" in norm_cid:
            target_cidades.append("anapolis")
          if "cidades polo" in norm_cid or "expansao" in norm_cid:
            target_estados.append("sao paulo")

          target_cidades = list(set(target_cidades))
          target_estados = list(set(target_estados))

          if target_cidades:
            sub = sub[sub["norm_cidade"].isin(target_cidades)]
          elif target_estados:
            sub = sub[sub["norm_estado"].isin(target_estados)]

          allowed_types = parse_types(tipos_input)
          if "terrenos" in norm_cons or "lotes" in norm_cons:
            if "Terreno" not in allowed_types:
              allowed_types.append("Terreno")

          if allowed_types:
            sub = sub[sub["Tipo de Bem"].isin(allowed_types)]

          min_v, max_v = parse_budget(valor_input)
          if any(
              p in norm_cons or p in norm_val
              for p in [
                  "qualquer valor",
                  "valores menores",
                  "100 a 500",
                  "todos os valores",
                  "qualquer",
              ]
          ):
            min_v, max_v = 0, 999999999

          if min_v > 0 and max_v < 999999999:
            sub = sub[
                (sub["preco_effective"] >= min_v)
                & (sub["preco_effective"] <= max_v)
            ]
          elif min_v > 0:
            sub = sub[sub["preco_effective"] >= min_v]
          elif max_v < 999999999:
            sub = sub[sub["preco_effective"] <= max_v]

          sub_sorted = sub.sort_values(
              by=["desconto_%", "preco_effective"], ascending=[False, True]
          )

          for _, imovel in sub_sorted.iterrows():
            resultados.append({
                "ID Investidor": idx + 1,
                "Nome do Investidor": nome,
                "Cidades Solicitadas": cidades_input,
                "Faixa Solicitada": valor_input,
                "Título do Imóvel": imovel["Título"],
                "Cidade Imóvel": imovel["Cidade"],
                "Estado Imóvel": imovel["Estado"],
                "Tipo de Bem": imovel["Tipo de Bem"],
                "Preço do Leilão (R$)": imovel["preco_effective"],
                "Valor de Avaliação (R$)": imovel[
                    "Valor de Avaliação do Leiloeiro"
                ],
                "Desconto (%)": round(imovel["desconto_%"], 2),
                "Endereço": imovel["Endereço"],
                "Link do Imóvel": imovel["Link"],
            })

        df_final = pd.DataFrame(resultados)

        output = io.BytesIO()
        with pd.ExcelWriter(output, engine="openpyxl") as writer:
          df_final.to_excel(writer, index=False)
        processed_data = output.getvalue()

        st.success(f"✅ Concluído! Total de {len(df_final):,} imóveis cruzados.")
        st.download_button(
            label="📥 Baixar Planilha Consolidada (.xlsx)",
            data=processed_data,
            file_name="cruzamento_leiloes_investidores.xlsx",
            mime=(
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            ),
        )
        st.dataframe(df_final.head(50))

      except Exception as e:
        st.error(f"Erro ao processar: {e}")
