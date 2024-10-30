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
nome_municipio = 'São Paulo'
sigla_uf = 'SP'
redes = 'Municipal'
etapas = 'Anos Iniciais'
anos = 2019

etapa = etapas[1]
ano = anos[1]

adp = adapter(db_path)
df = adp %>% adapter.fetch_pca_data_schools(
  nome_municipio = nome_municipio,
  sigla_uf = sigla_uf,
  redes = redes,
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

doc = doc %>% ppt.add_transition(
  text = paste0(
    "ACP das escolas da rede ",rede," dos ",
    tolower(etapa)," do ensino fundamental do município de ",
    nome_municipio," - ",sigla_uf
  )
)

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
  position = 'center'
)


doc %>% ppt.save(output_path)
