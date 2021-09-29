#head_dept<-read_csv("dados/head_dept.csv", col_names = FALSE)
#titulo = 'LISTA DE ALUNOS EGRESSOS DAS ESCOLAS ESTADUAIS'
#colunas <-read_csv2('dados/lista_colunas.csv', col_names = FALSE)
#fonte<-"Fonte: SEDUC/SIGE Escola em 26/09/2021"

#df=relatorio %>% filter(nr_anoletivo==ano) %>% 
#  rename(id_crede =id_crede_sefor,
#         nr_ddd_res	=	nr_ddd_residencia_responsavel,
#         nr_fone_re	=	nr_fone_residencia_responsavel,
#         nr_ddd_cel	=	nr_ddd_celular_responsavel,
#         nr_fone_ce	=	nr_fone_celular_responsavel
#  )
#nm_sheet = 'RELATORIO'
#nr_sheet = 1
ncol=ncol(df)
nrow =nrow(df)+15

addWorksheet(wb, nm_sheet)
writeData(wb, nr_sheet,head_dept,xy = c("A",7), rowNames = F, colNames = F)
writeData(wb, nr_sheet, titulo,xy = c("A",11), rowNames = F, colNames = F)
writeData(wb, nr_sheet, fonte,xy = c("A",12), rowNames = F, colNames = F)
writeData(wb, nr_sheet, colunas,xy = c("A",14), rowNames = F, colNames = F)
writeData(wb, nr_sheet, df,xy = c("A",15),rowNames = F, colNames = TRUE,withFilter = TRUE)

addStyle(wb, nr_sheet, header_style, rows = 1:11, cols = 1:ncol, gridExpand = TRUE)
addStyle(wb, nr_sheet, bodyStyle, rows = 16:nrow, cols = 1:ncol, gridExpand = TRUE)
addStyle(wb, nr_sheet, filterStyle, rows = 15, cols = 1:ncol, gridExpand = TRUE)
addStyle(wb, nr_sheet, colStyle, rows = 14, cols = 1:ncol, gridExpand = TRUE)
setRowHeights(wb, sheet = 1, row = 14, heights = 30)
#freezePane(wb, nr_sheet, firstRow = TRUE)
setColWidths(wb, nr_sheet, cols = c(1,2), widths = 10)
#setColWidths(wb, nr_sheet, cols = 2, widths = 15)
setColWidths(wb, nr_sheet, cols = 2:ncol, widths = "auto")
insertImage( wb,  nr_sheet,'img/002_marca_horizontal_color.png', startCol = 1,
             startRow = 1, width = 3.128, height = 1.2155)

#saveWorkbook(wb, file = paste0("~/Temporarios/teste.xlsx"), overwrite = T)

