library(lcde.toolbox)
library(lcde.client)
library(tcltk)

# janela = tktoplevel()
# tkwm.title(janela, "Relatório PCA")
#
# opcoes = c("Opção 1", "Opção 2", "Opção 3")
#
# lista = tklistbox(janela, height = length(opcoes), selectmode = "single")
# for (opcao in opcoes) {
#   tkinsert(lista, "end", opcao)
# }
# tkgrid(lista)
#
# botao_funcao <- function() {
#   indice <- as.integer(tkcurselection(lista)) + 1  # Converter índice
#   opcao_selecionada <- opcoes[indice]
#   tkmessageBox(title = "Seleção", message = paste("Você selecionou:", opcao_selecionada))
# }
#
# botao <- tkbutton(janela, text = "Confirmar Seleção", command = botao_funcao)
# tkgrid(botao)

db_path = 'C:/Users/pedro/Documents/iea/banco_pca.sqlite3'
template_path = 'D:/IEA/report_pca/template.pptx'
output_path = 'C:/Users/pedro/Desktop/report.pptx'
nome_municipio = 'Ribeirão Preto'
sigla_uf = 'SP'
rede = 'Municipal'
etapas = 'Anos Iniciais'
anos = 2019
add_boundary = FALSE
add_surface = FALSE
ano_inse = 2021

etapa = etapas[1]
ano = anos[1]

adp = adapter(db_path)
df = adp %>% adapter.fetch_pca_data_schools(
  nome_municipio = nome_municipio,
  sigla_uf = sigla_uf,
  redes = rede,
  etapas = etapas,
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
    " – 2019 e 2023"
  )
)

doc = doc %>% ppt.on_slide(2)

# Transição de etapa
doc = doc %>% ppt.add_transition(
  text = paste0(
    "ACP das escolas da rede ",rede," dos ",
    tolower(etapa)," do ensino fundamental do município de ",
    nome_municipio," - ",sigla_uf
  )
)

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
df_table = if(nrow(df_table > 24)) df_table[1:24,] else df_table

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

doc %>% ppt.save(output_path)
