#Script controlador das planilhas xlsx
library(openxlsx)


## Criação do dataframe ----

# Leitura de arquivo .csv
#library(readr)
#df<-read_delim("~/arquyivo.csv", delim = ";", 
#           escape_double = FALSE, 
#           locale = locale(encoding = "ISO-8859-1"), 
#           trim_ws = TRUE)

# Leitura de banco de dados

#source('../conexoes/dbi_astim_seduc.R')
#source('sql/query_generica.R')
#df<-dbGetQuery(con,query)


## Configurações da planilha: ----
for (variable in c(1)) {
library(readr)
head_dept<-read_csv("dados/head_dept.csv", col_names = FALSE)
colunas<-read_csv2('dados/lista_colunas_escola.csv',col_names = FALSE) # Para dados por escola
#colunas$X8 = "NOME DA ELETIVA"
#colunas$X9 = "CÓDIGO"
#colunas$X10 = "SEMESTRE"
#colunas$X11 = "QTD. TURMAS"
#colunas$X12 = "QTD. ALUNOS"

  
titulo = 'LISTA DE ESCOLAS EEMTI QUE OFERTAM A DISCIPLINA EDUCAÇÃO PARA CIDADANIA NA ESCOLA'

fonte<-"Fonte: SEDUC/SIGE Escola em 29/09/2021"

nm_sheet = 'RELATORIO'
nr_sheet = 1
local_nome_arquivo = "~/Temporarios/eletivas_eemti_2021_2.xlsx"


  wb <- createWorkbook()
  source('script_xlsx_style.R')
  source('script_xlsx_body.R')
  saveWorkbook(wb, file = paste0(local_nome_arquivo), overwrite = T)
  
#CLIQUE NO LADO DIREITO DO FECHEMANTO DE CHAVE ABAIXO E DEPOIS  CTRL + ENTER
  }
#