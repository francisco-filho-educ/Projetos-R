library(openxlsx)
header_style <- createStyle(halign = "left", textDecoration = "bold")
font_style <- createStyle(halign = "left", textDecoration = "italic")
bodyStyle <- createStyle(border = "TopBottomLeftRight")
colStyle <- createStyle(border = "TopBottomLeftRight",halign = "center", 
                        textDecoration = "bold" ,fgFill = "#3CB371",wrapText=TRUE)
filterStyle <- createStyle(fontColour = '#FFFFFF',border = "TopBottomLeftRight")