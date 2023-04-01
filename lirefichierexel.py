import openpyxl
readbook1=openpyxl.load_workbook("decembre.xlsx",data_only=True)
readbook2=openpyxl.load_workbook("octobre.xlsx",data_only=True)
readbook3=openpyxl.load_workbook("novembre.xlsx",data_only=True)

feuille1=readbook2.active
donne={}

for i in range(2,feuille1.max_row):
    cle_donne=feuille1.cell(i,1).value
    valeur_donne=feuille1.cell(i,4).value
    donne[cle_donne]=valeur_donne
    if not cle_donne:
      break
    
print(donne)