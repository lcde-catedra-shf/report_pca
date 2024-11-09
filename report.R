library(lcde.toolbox)
library(lcde.client)

db_path = 'C:/Users/pedro/Documents/iea/banco_pca.sqlite3'
template_path = 'C:/Users/pedro/Documents/iea/report_pca/template.pptx'
output_path = 'C:/Users/pedro/Desktop/report.pptx'
nome_municipio = 'Ribeirão Preto'
sigla_uf = 'SP'
rede = 'Municipal'
etapas = c('Anos Iniciais', 'Anos Finais')
anos = c(2019, 2023)
add_boundary = TRUE
add_surface = TRUE
ano_inse = 2019

adp = adapter(db_path)

# doc = ppt.from_template(template_path)
# doc = doc %>% ppt.on_slide(2)
# doc = doc %>% ppt.new_slide(title = 'Gráficos de Radar')
#
# # Obtendo dados do Estado
# dados_estado = adp %>% adapter.fetch_state('SP')
#
# radar_data = adp %>% adapter.fetch_pca_data_municipality(
#   nome_municipio = 'Ribeirão Preto',
#   sigla_uf = 'SP',
#   rede = 'Pública',
#   etapa = 'Anos Iniciais',
#   localizacao = 'Total',
#   ano = 2019
# )
#
# # Gráficos de Radar
# for(i in 1:length(anos)) {
#   for(j in 1:3) {
#
#     if(j == 1) {
#       radar_data = adp %>% adapter.fetch_pca_data_municipality(
#         nome_municipio = nome_municipio,
#         sigla_uf = sigla_uf,
#         rede = 'Pública',
#         etapa = 'Anos Iniciais',
#         localizacao = 'Total',
#         anos[i]
#       )
#       radar_title = paste0(nome_municipio,', ',sigla_uf, ' - ', anos[i])
#       radar_color = colors.mixed()[1]
#     } else if(j == 2) {
#       radar_data = adp %>% adapter.fetch_pca_data_state(
#         sigla_uf = sigla_uf,
#         rede = 'Pública',
#         etapa = 'Anos Iniciais',
#         localizacao = 'Total',
#         ano = anos[i]
#       )
#       radar_title = paste0(dados_estado$nome_estado, ' - ', anos[i])
#       radar_color = colors.mixed()[6]
#     } else {
#       radar_data = adp %>% adapter.fetch_pca_data_country(
#         nome_pais = 'Brasil',
#         rede = 'Pública',
#         etapa = 'Anos Iniciais',
#         localizacao = 'Total',
#         ano = anos[i]
#       )
#       radar_title = paste0(dados_estado$nome_pais, ' - ', anos[i])
#       radar_color = colors.mixed()[5]
#     }
#
#     if(nrow(radar_data) == 0) {
#       next()
#     }
#
#     radar_data = na.omit(radar_data)
#     radar_data$mat = round(radar_data$mat, 0)
#     radar_data$lp = round(radar_data$lp, 0)
#     radar_data$tdi = round(radar_data$tdi, 0)
#     radar_data$np = round(radar_data$np * 10, 1)
#     radar_data$fluxo = round(radar_data$fluxo*100, 0)
#
#     radar_data = radar_data[,c('lp', 'mat', 'np', 'fluxo', 'tdi')]
#     colnames(radar_data) = c('LP', 'MAT', 'NP', 'FLUXO', 'DIS')
#
#     radar_plot = ggviz.radar(
#       data = radar_data,
#       colors = radar_color,
#       show_score = TRUE,
#       title = radar_title,
#       size = 'normal'
#     )
#
#     doc = doc %>% ppt.add_ggplot(
#       ggplot_obj = radar_plot,
#       position = pptpos.grid(
#         n_rows = length(anos),
#         n_columns = 4,
#         row = i,
#         column = j,
#         margin = 0.01
#       )
#     )
#   }
# }
#
# doc %>% ppt.save(output_path)


for(etapa in etapas) {
  for(ano in anos) {
    df = adp %>% adapter.fetch_pca_data_schools(
      nome_municipio = nome_municipio,
      sigla_uf = sigla_uf,
      redes = rede,
      etapas = etapa,
      anos = ano
    )

    if(nrow(df) == 0) {
      stop(paste0(
        "Não existem dados dos indicadores para ", ano, ", ", etapa
      ))
    }
  }
}

if(add_surface) {
  inse = adp %>% adapter.fetch_municipality_inse(
    nome_municipio = nome_municipio,
    sigla_uf = sigla_uf,
    rede = rede,
    ano = ano_inse
  )

  if(nrow(inse) == 0) {
    stop(paste0(
      "Não existem dados de nível socioeconômico para ", ano_inse
    ))
  }
}

georef_obj = if(add_surface || add_boundary) {
  georef.from_geojson(
    adp %>% adapter.fetch_municipality_boundary(
      nome_municipio = nome_municipio,
      sigla_uf = sigla_uf
    ) %>% .$malha
  )
} else {
  NULL
}

surface_data = NULL
surface_latitude = NULL
surface_longitude = NULL

if(add_surface) {
  inse = adp %>% adapter.fetch_municipality_inse(
    nome_municipio = nome_municipio,
    sigla_uf = sigla_uf,
    rede = rede,
    ano = ano_inse
  )

  surface_data = inse$inse
  surface_latitude = inse$latitude
  surface_longitude = inse$longitude
}

# Título
doc = ppt.from_template(template_path) %>%
  ppt.on_slide(1)

doc = doc %>% ppt.add_document_title(
  title = paste0(
    "ACP das Escolas da rede ",rede,
    " do Município de ",nome_municipio,
    " – ", anos[1] ," a ", anos[length(anos)]
  )
)

doc = doc %>% ppt.on_slide(2)

for(etapa in etapas) {
  for(ano in anos) {
    # Transição de ano e etapa
    doc = doc %>% ppt.add_transition(
      text = paste0(
        "Análise de Componentes Principais\n",
        "Rede ",rede, "\n",
        etapa," do Ensino Fundamental\n",
        ano
      )
    )

    df = adp %>% adapter.fetch_pca_data_schools(
      nome_municipio = nome_municipio,
      sigla_uf = sigla_uf,
      redes = rede,
      etapas = etapa,
      anos = ano
    )
    df = na.omit(df)
    df$mat = round(df$mat, 0)
    df$lp = round(df$lp, 0)
    df$tdi = round(df$tdi, 0)
    df$np = round(df$np * 10, 1)
    df$fluxo = round(df$fluxo*100, 0)

    dfpca = df[,c('lp', 'mat', 'np', 'fluxo', 'tdi')]
    colnames(dfpca) = c('LP', 'MAT', 'NP', 'FLUXO', 'DIS')
    pca_obj = pca.from_data_frame(dfpca)

    # Matriz de indicadores
    df_table = data.frame(
      'Nome da escola' = df$nome_escola,
      'Nome abreviado' = inep.abbreviate_school_names(df$nome_escola),
      'LP' = df$lp,
      'MAT' = df$mat,
      'NP' = df$np,
      'FLUXO' = df$fluxo,
      'DIS' = df$tdi
    )
    df_table = if(nrow(df_table) > 24) df_table[1:24,] else df_table

    matriz_indicadores = table(
      df_table,
      column_names = c(
        'Nome da escola', 'Nome abreviado',
        'LP', 'MAT', 'NP', 'FLUXO', 'DIS'
      )
    )

    doc = doc %>% ppt.new_slide(
      title = paste0(
        "Matriz de indicadores - ",etapa,", ", ano
      )
    )

    doc = doc %>% ppt.add_table(
      table_obj = matriz_indicadores,
      position = 'center-large'
    )

    # Dispersão
    doc = doc %>% ppt.new_slide(
      title = paste0("ACP da Rede ",rede," – ",etapa,", ", ano)
    )

    doc = doc %>% ppt.add_ggplot(
      ggplot_obj = pcaviz.scatter(
        pca_obj = pca_obj,
        labels = inep.abbreviate_school_names(df$nome_escola)
      ),
      position = 'center'
    )

    # Variância explicada
    doc = doc %>% ppt.new_slide(
      title = paste0(
        "Percentual de variância explicada - ",
        rede," – ",etapa,", ", ano
      )
    )

    doc = doc %>% ppt.add_ggplot(
      ggplot_obj = pcaviz.explained_variance(
        pca_obj = pca_obj
      ),
      position = 'center'
    )

    # Cargas
    doc = doc %>% ppt.new_slide(
      title = paste0(
        "Cargas dos componentes principais – ",
        rede," – ",etapa,", ", ano
      )
    )

    doc = doc %>% ppt.add_ggplot(
      ggplot_obj = pcaviz.component_loads(
        pca_obj = pca_obj,
        component = 1
      ) %>% pcaviz.add_title("Componente 1"),
      position = 'left-half'
    )

    doc = doc %>% ppt.add_ggplot(
      ggplot_obj = pcaviz.component_loads(
        pca_obj = pca_obj,
        component = 2
      ) %>% pcaviz.add_title("Componente 2"),
      position = 'right-half'
    )

    # Mapas
    labels = if(nrow(df) > 30) {
      NULL
    } else {
      labels = inep.abbreviate_school_names(df$nome_escola)
    }

    # Mapa matemática
    doc = doc %>% ppt.new_slide(
      title = paste0(
        "Aprendizado adequado em matemática – ",
        rede," – ",etapa,", ", ano
      )
    )

    doc = doc %>% ppt.add_ggplot(
      ggplot_obj = geogg.percentage_of_proficiency_map(
        data = df$mat,
        subject = 'mathematics',
        latitude = df$latitude,
        longitude = df$longitude,
        labels = labels,
        add_surface = add_surface,
        add_boundary = add_boundary,
        georef_obj = georef_obj,
        surface_data = surface_data,
        surface_latitude = surface_latitude,
        surface_longitude = surface_longitude,
        surface_legend = 'Nível Socioeconômico'
      ),
      position = 'center'
    )

    # Mapa matemática
    doc = doc %>% ppt.new_slide(
      title = paste0(
        "Aprendizado adequado em língua portuguesa – ",
        rede," – ",etapa,", ", ano
      )
    )

    doc = doc %>% ppt.add_ggplot(
      ggplot_obj = geogg.percentage_of_proficiency_map(
        data = df$lp,
        subject = 'portuguese language',
        latitude = df$latitude,
        longitude = df$longitude,
        labels = labels,
        add_surface = add_surface,
        add_boundary = add_boundary,
        georef_obj = georef_obj,
        surface_data = surface_data,
        surface_latitude = surface_latitude,
        surface_longitude = surface_longitude,
        surface_legend = 'Nível Socioeconômico'
      ),
      position = 'center'
    )

    # Mapa ACP
    doc = doc %>% ppt.new_slide(
      title = paste0(
        "Desempenho relativo – ",
        rede," – ",etapa,", ", ano
      )
    )

    doc = doc %>% ppt.add_ggplot(
      ggplot_obj = geogg.pca_map(
        pca_obj = pca_obj,
        latitude = df$latitude,
        longitude = df$longitude,
        labels = labels,
        add_surface = add_surface,
        add_boundary = add_boundary,
        georef_obj = georef_obj,
        surface_data = surface_data,
        surface_latitude = surface_latitude,
        surface_longitude = surface_longitude,
        surface_legend = 'Nível Socioeconômico'
      ),
      position = 'center'
    )
  }


  # ACP com todos os anos
  df = adp %>% adapter.fetch_pca_data_schools(
    nome_municipio = nome_municipio,
    sigla_uf = sigla_uf,
    redes = rede,
    etapas = etapa,
    anos = anos
  )
  df = na.omit(df)
  df$mat = round(df$mat, 0)
  df$lp = round(df$lp, 0)
  df$tdi = round(df$tdi, 0)
  df$np = round(df$np * 10, 1)
  df$fluxo = round(df$fluxo*100, 0)

  dfpca = df[,c('lp', 'mat', 'np', 'fluxo', 'tdi')]
  colnames(dfpca) = c('LP', 'MAT', 'NP', 'FLUXO', 'DIS')
  pca_obj = pca.from_data_frame(dfpca)

  # Transição de ano e etapa
  doc = doc %>% ppt.add_transition(
    text = paste0(
      "Análise de Componentes Principais\n",
      "Rede ",rede, "\n",
      etapa," do Ensino Fundamental\n",
      anos[1], ' a ', anos[length(anos)]
    )
  )

  # Dispersão
  doc = doc %>% ppt.new_slide(
    title = paste0(
      "ACP da Rede ",
      rede," – ",etapa,", ", anos[1], " a ", anos[length(anos)]
    )
  )

  doc = doc %>% ppt.add_ggplot(
    ggplot_obj = pcaviz.scatter(
      pca_obj = pca_obj,
      groups = as.factor(df$ano),
      include_ID = FALSE
    ),
    position = 'center'
  )

  # Variância explicada
  doc = doc %>% ppt.new_slide(
    title = paste0(
      "Percentual de variância explicada - ",
      rede," – ",etapa,", ", anos[1], " a ", anos[length(anos)]
    )
  )

  doc = doc %>% ppt.add_ggplot(
    ggplot_obj = pcaviz.explained_variance(
      pca_obj = pca_obj
    ),
    position = 'center'
  )

  # Cargas
  doc = doc %>% ppt.new_slide(
    title = paste0(
      "Cargas dos componentes principais – ",
      rede," – ",etapa,", ", anos[1], " a ", anos[length(anos)]
    )
  )

  doc = doc %>% ppt.add_ggplot(
    ggplot_obj = pcaviz.component_loads(
      pca_obj = pca_obj,
      component = 1
    ) %>% pcaviz.add_title("Componente 1"),
    position = 'left-half'
  )

  doc = doc %>% ppt.add_ggplot(
    ggplot_obj = pcaviz.component_loads(
      pca_obj = pca_obj,
      component = 2
    ) %>% pcaviz.add_title("Componente 2"),
    position = 'right-half'
  )


  # Maiores Variações
  doc = doc %>% ppt.new_slide(
    title = paste0(
      "Maiores variações de desempenho – ",
      rede," – ",etapa,", ", anos[1], " a ", anos[length(anos)]
    )
  )

  doc = doc %>% ppt.add_ggplot(
    ggplot_obj = pcaviz.scatter(
      pca_obj = pca_obj,
      groups = as.factor(df$ano),
      include_ID = FALSE
    ) %>% pcaviz.add_largest_variations(
      number = 3,
      keys = df$id,
      years = df$ano,
      labels = inep.abbreviate_school_names(df$nome_escola),
      variation = 'positive'
    ),
    position = 'wide-right'
  )

  for(i in 1:3) {
    doc = doc %>% ppt.add_table(
      table_obj = table.pca_variation(
        pca_obj,
        rank = i,
        keys = df$id,
        years = df$ano,
        labels = inep.abbreviate_school_names(df$nome_escola),
        variation = 'positive'
      ),
      position = pptpos.grid(3,4,i,1),
      fit_height = TRUE
    )
  }

  # Menores Variações
  doc = doc %>% ppt.new_slide(
    title = paste0(
      "Menores variações de desempenho – ",
      rede," – ",etapa,", ", anos[1], " a ", anos[length(anos)]
    )
  )

  doc = doc %>% ppt.add_ggplot(
    ggplot_obj = pcaviz.scatter(
      pca_obj = pca_obj,
      groups = as.factor(df$ano),
      include_ID = FALSE
    ) %>% pcaviz.add_largest_variations(
      number = 3,
      keys = df$id,
      years = df$ano,
      labels = inep.abbreviate_school_names(df$nome_escola),
      variation = 'negative'
    ),
    position = 'wide-right'
  )

  for(i in 1:3) {
    doc = doc %>% ppt.add_table(
      table_obj = table.pca_variation(
        pca_obj,
        rank = i,
        keys = df$id,
        years = df$ano,
        labels = inep.abbreviate_school_names(df$nome_escola),
        variation = 'negative'
      ),
      position = pptpos.grid(3,4,i,1),
      fit_height = TRUE
    )
  }

  # Mapa MAT
  doc = doc %>% ppt.new_slide(
    title = paste0(
      "Aprendizado em matemática – ",
      rede," – ",etapa,", ", anos[1], " a ", anos[length(anos)]
    )
  )

  for(i in 1:length(anos)) {
    ano = anos[i]

    df = adp %>% adapter.fetch_pca_data_schools(
      nome_municipio = nome_municipio,
      sigla_uf = sigla_uf,
      redes = rede,
      etapas = etapa,
      anos = ano
    )
    df = na.omit(df)
    df$mat = round(df$mat, 0)
    df$lp = round(df$lp, 0)
    df$tdi = round(df$tdi, 0)
    df$np = round(df$np * 10, 1)
    df$fluxo = round(df$fluxo*100, 0)

    dfpca = df[,c('lp', 'mat', 'np', 'fluxo', 'tdi')]
    colnames(dfpca) = c('LP', 'MAT', 'NP', 'FLUXO', 'DIS')
    pca_obj = pca.from_data_frame(dfpca)

    doc = doc %>% ppt.add_ggplot(
      ggplot_obj = geogg.percentage_of_proficiency_map(
        data = df$mat,
        subject = 'mathematics',
        latitude = df$latitude,
        longitude = df$longitude,
        add_surface = add_surface,
        add_boundary = add_boundary,
        georef_obj = georef_obj,
        surface_data = surface_data,
        surface_latitude = surface_latitude,
        surface_longitude = surface_longitude,
        surface_legend = 'Nível Socioeconômico'
      ) %>%
        geogg.without_legend() %>%
        geogg.add_title(paste0(ano)),
      position = pptpos.grid(1, length(anos), 1, i, margin=0.03)
    )
  }

  # Mapa LP
  doc = doc %>% ppt.new_slide(
    title = paste0(
      "Aprendizado em Língua Portuguesa – ",
      rede," – ",etapa,", ", anos[1], " a ", anos[length(anos)]
    )
  )

  for(i in 1:length(anos)) {
    ano = anos[i]

    df = adp %>% adapter.fetch_pca_data_schools(
      nome_municipio = nome_municipio,
      sigla_uf = sigla_uf,
      redes = rede,
      etapas = etapa,
      anos = ano
    )
    df = na.omit(df)
    df$mat = round(df$mat, 0)
    df$lp = round(df$lp, 0)
    df$tdi = round(df$tdi, 0)
    df$np = round(df$np * 10, 1)
    df$fluxo = round(df$fluxo*100, 0)

    dfpca = df[,c('lp', 'mat', 'np', 'fluxo', 'tdi')]
    colnames(dfpca) = c('LP', 'MAT', 'NP', 'FLUXO', 'DIS')
    pca_obj = pca.from_data_frame(dfpca)

    doc = doc %>% ppt.add_ggplot(
      ggplot_obj = geogg.percentage_of_proficiency_map(
        data = df$lp,
        subject = 'portuguese language',
        latitude = df$latitude,
        longitude = df$longitude,
        add_surface = add_surface,
        add_boundary = add_boundary,
        georef_obj = georef_obj,
        surface_data = surface_data,
        surface_latitude = surface_latitude,
        surface_longitude = surface_longitude,
        surface_legend = 'Nível Socioeconômico'
      ) %>%
        geogg.without_legend() %>%
        geogg.add_title(paste0(ano)),
      position = pptpos.grid(1, length(anos), 1, i, margin=0.03)
    )
  }

  # Mapa ACP
  doc = doc %>% ppt.new_slide(
    title = paste0(
      "Desempenho relativo – ",
      rede," – ",etapa,", ", anos[1], " a ", anos[length(anos)]
    )
  )

  for(i in 1:length(anos)) {
    ano = anos[i]

    df = adp %>% adapter.fetch_pca_data_schools(
      nome_municipio = nome_municipio,
      sigla_uf = sigla_uf,
      redes = rede,
      etapas = etapa,
      anos = ano
    )
    df = na.omit(df)
    df$mat = round(df$mat, 0)
    df$lp = round(df$lp, 0)
    df$tdi = round(df$tdi, 0)
    df$np = round(df$np * 10, 1)
    df$fluxo = round(df$fluxo*100, 0)

    dfpca = df[,c('lp', 'mat', 'np', 'fluxo', 'tdi')]
    colnames(dfpca) = c('LP', 'MAT', 'NP', 'FLUXO', 'DIS')
    pca_obj = pca.from_data_frame(dfpca)

    doc = doc %>% ppt.add_ggplot(
      ggplot_obj = geogg.pca_map(
        pca_obj = pca_obj,
        latitude = df$latitude,
        longitude = df$longitude,
        add_surface = add_surface,
        add_boundary = add_boundary,
        georef_obj = georef_obj,
        surface_data = surface_data,
        surface_latitude = surface_latitude,
        surface_longitude = surface_longitude,
        surface_legend = 'Nível Socioeconômico'
      ) %>%
        geogg.without_legend() %>%
        geogg.add_title(paste0(ano)),
      position = pptpos.grid(1, length(anos), 1, i, margin=0.03)
    )
  }
}

doc %>% ppt.save(output_path)


