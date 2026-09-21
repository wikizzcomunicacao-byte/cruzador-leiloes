if file_leiloes and file_investidores and executar:
    with st.spinner("A analisar critérios e a cruzar bases..."):
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
            & (df_leiloes["Valor de Avaliação do Leiloeiro"] > 0)
            & (pd.notnull(df_leiloes["preco_effective"]))
            & (df_leiloes["preco_effective"] > 0),
            (
                (
                    df_leiloes["Valor de Avaliação do Leiloeiro"]
                    - df_leiloes["preco_effective"]
                )
                / df_leiloes["Valor de Avaliação do Leiloeiro"]
            )
            * 100,
            0,
        )

        cols = list(df_investidores.columns)
        resultados = []
        investidores_sem_imoveis = 0

        # O loop DEVE estar indentado aqui dentro!
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

          if cidades_input and cidades_input.lower() != "nan":
            partes = re.split(r"[,/;\n]", cidades_input)
            for parte in partes:
              p_norm = normalize(parte)
              if p_norm:
                target_cidades.append(p_norm)

          if any(
              w in norm_cid or w in norm_cons
              for w in [
                  "estado de sp",
                  "sao paulo (estado)",
                  "interior de sao paulo",
              ]
          ):
            target_estados.append("sao paulo")

          target_cidades = list(set(target_cidades))
          target_estados = list(set(target_estados))

          if target_cidades:
            sub = sub[
                sub["norm_cidade"].apply(
                    lambda x: any(tc in x for tc in target_cidades)
                )
            ]
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

          if sub_sorted.empty:
            investidores_sem_imoveis += 1

          for _, imovel in sub_sorted.iterrows():
            preco = imovel["preco_effective"]
            avaliac = imovel["Valor de Avaliação do Leiloeiro"]

            valor_referencia = avaliac
            custos_adicionais = (preco * taxa_leiloeiro) + (preco * taxa_itbi)
            custo_total = preco + custos_adicionais
            lucro_liquido = (
                valor_referencia - custo_total
                if pd.notnull(valor_referencia) and pd.notnull(preco)
                else 0
            )

            resultados.append({
                "ID Investidor": idx + 1,
                "Nome do Investidor": nome,
                "Cidades Solicitadas": cidades_input,
                "Faixa Solicitada": valor_input,
                "Título do Imóvel": imovel["Título"],
                "Cidade Imóvel": imovel["Cidade"],
                "Estado Imóvel": imovel["Estado"],
                "Tipo de Bem": imovel["Tipo de Bem"],
                "Preço do Leilão (R$)": preco,
                "Valor de Avaliação (R$)": avaliac,
                "Desconto (%)": round(imovel["desconto_%"], 2),
                "Custo Total Estimado (R$)": round(custo_total, 2),
                "Lucro Líquido Real (R$)": round(lucro_liquido, 2),
                "Endereço": imovel["Endereço"],
                "Link do Imóvel": imovel["Link"],
            })

        st.session_state["df_final"] = pd.DataFrame(resultados)
        st.session_state["investidores_sem_imoveis"] = investidores_sem_imoveis
        st.session_state["imoveis_selecionados"] = []
        st.toast("✅ Cruzamento concluído com sucesso!", icon="🎉")

      except Exception as e:
        st.error(f"Erro ao processar as planilhas: {e}")
