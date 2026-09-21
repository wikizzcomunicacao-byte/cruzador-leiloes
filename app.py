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

          # -------------------------------------------------------------
          # ABORDAGEM DINÂMICA: Extração inteligente das cidades informadas
          # -------------------------------------------------------------
          if cidades_input and cidades_input.lower() != "nan":
            # Separa por vírgula, barra ou ponto e vírgula caso o investidor digite várias cidades
            partes = re.split(r"[,/;\n]", cidades_input)
            for parte in partes:
              p_norm = normalize(parte)
              if p_norm:
                target_cidades.append(p_norm)

          # Se houver menção genérica a estado de SP ou termos amplos nas considerações
          if any(w in norm_cid or w in norm_cons for w in ["estado de sp", "sao paulo (estado)", "interior de sao paulo"]):
            target_estados.append("sao paulo")

          target_cidades = list(set(target_cidades))
          target_estados = list(set(target_estados))

          # -------------------------------------------------------------
          # FILTRAGEM AUTOMÁTICA NA BASE DE LEILÕES
          # -------------------------------------------------------------
          if target_cidades:
            # Filtra onde a cidade normalizada do leilão contém ou coincide com as cidades pedidas
            sub = sub[sub["norm_cidade"].apply(lambda x: any(tc in x for tc in target_cidades))]
          elif target_estados:
            sub = sub[sub["norm_estado"].isin(target_estados)]

          # Filtro de Tipos de Imóvel
          allowed_types = parse_types(tipos_input)
          if "terrenos" in norm_cons or "lotes" in norm_cons:
            if "Terreno" not in allowed_types:
              allowed_types.append("Terreno")

          if allowed_types:
            sub = sub[sub["Tipo de Bem"].isin(allowed_types)]

          # Filtro de Orçamento
          min_v, max_v = parse_budget(valor_input)
          if any(p in norm_cons or p in norm_val for p in ["qualquer valor", "valores menores", "100 a 500", "todos os valores", "qualquer"]):
            min_v, max_v = 0, 999999999

          if min_v > 0 and max_v < 999999999:
            sub = sub[(sub["preco_effective"] >= min_v) & (sub["preco_effective"] <= max_v)]
          elif min_v > 0:
            sub = sub[sub["preco_effective"] >= min_v]
          elif max_v < 999999999:
            sub = sub[sub["preco_effective"] <= max_v]

          sub_sorted = sub.sort_values(by=["desconto_%", "preco_effective"], ascending=[False, True])

          if sub_sorted.empty:
            investidores_sem_imoveis += 1

          for _, imovel in sub_sorted.iterrows():
            preco = imovel["preco_effective"]
            avaliac = imovel["Valor de Avaliação do Leiloeiro"]
            
            valor_referencia = avaliac
            custos_adicionais = (preco * taxa_leiloeiro) + (preco * taxa_itbi)
            custo_total = preco + custos_adicionais
            lucro_liquido = valor_referencia - custo_total if pd.notnull(valor_referencia) and pd.notnull(preco) else 0

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
