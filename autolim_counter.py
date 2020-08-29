import xlrd
import os
import sys

current_directory = os.getcwd()
#print (current_directory)

filePath = current_directory + "/ventas.xlsx"
#filePath = "/home/nahuel/Documents/autolim/ventas.xlsx"

openFile = xlrd.open_workbook(filePath)

sheet = openFile.sheet_by_name("Ventas AR")

mes  = input('ingrese el mes que desea analizar: ')

common_sales = []
full_sales = []
total_full_sales = {'autolim':0, 'micro':0, 'xl':0, 'cepillo':0 }
total_common_sales = {'autolim':0, 'micro':0, 'xl':0, 'cepillo':0 }
column_date = 1
column_shipping = 25
column_id = 14
column_quantity = 5
x1_autolim = ["Limpia Tu Tapizado Autolim Telas Cueros Alfombra Plasticos","Limpia Tu Tapizado  Autolim Telas Cueros Alfombra Plasticos","Limpiador Sillones Chenille Pana Tela Gamuza Cuero Autolim","Crema Limpia Tapizados Telas Cuero Butacas Detailing Autolim","Limpia Tapizados Telas Cuero Butacas Techo Detailing Autolim","Limpiador Tapizados Sillones Chenille Pana Tela Autolim",]
no_match = []
# print("Num de filas:", sheet.nrows ) # example to print number of rows

for i in range(sheet.nrows - 1 ):
    cellValue = sheet.cell_value(i,column_date)
    #print (mes in str(cellValue))
    if mes in str(cellValue): # okDate
        #if str(cellValue).find(mes): # okDate   
        cellValue_shipping = sheet.cell_value(i,column_shipping)
        cellValue_id = sheet.cell_value(i,column_id)
        cellValue_quantity = sheet.cell_value(i,column_quantity)
        if "Full" in cellValue_shipping : # okFull
            full_sales.append((cellValue_id,cellValue_quantity))
        else:
            common_sales.append((cellValue_id,cellValue_quantity))

# calculate total common sales
for i in range(len(common_sales)):
    sale_id = common_sales[i][0]
    sale_quantity = common_sales[i][1]

    if sale_id in x1_autolim:
        total_common_sales["autolim"] += int(sale_quantity)

    elif sale_id == "Kit Autolim - Limpia Tapizados + Paño De Microfibra":
        total_common_sales["autolim"] += int(sale_quantity)
        total_common_sales["micro"]   += int(sale_quantity)

    elif sale_id == "Pack X 3 Limpia Tapizados Telas Cuero Butacas Techo Autolim":
        total_common_sales["autolim"] += 3 * int(sale_quantity)
        
    elif sale_id == "Kit Autolim X 2 Limpia Tapizados + 2 Paños De Microfibra":
        total_common_sales["autolim"] += 2 * int(sale_quantity)
        total_common_sales["micro"]   += 2 * int(sale_quantity)
    
    elif sale_id == "Limpiador Multiproposito + 2 Paños De Microfibra Autolim !":
        total_common_sales["autolim"] += int(sale_quantity)
        total_common_sales["micro"]   += 2 * int(sale_quantity)

    elif sale_id == "Kit Autolim - 2 Limpia Tapizados + 2 Paños De Microfibra Xl":
        total_common_sales["autolim"] += 2 * int(sale_quantity)
        total_common_sales["xl"]      += 2 * int(sale_quantity)

    elif sale_id == "Paño De Microfibra Autolim 37,5 X 37,5 Cm Limpieza Interior!":
        total_common_sales["micro"]   += int(sale_quantity)

    elif sale_id == "Paño De Microfibra Tamaño Xl Autolim - 61 X 77 Cm !!!":
        total_common_sales["xl"]   += int(sale_quantity)

    else: 
        no_match.append(common_sales[i])


# calculate total common sales
for i in range(len(full_sales)):
    sale_id = full_sales[i][0]
    sale_quantity = full_sales[i][1]

    if str(sale_id) in x1_autolim:
        total_full_sales["autolim"] += int(sale_quantity)
    #elif sale_id == "Limpia Tu Tapizado  Autolim Telas Cueros Alfombra Plasticos":
    #    total_full_sales["autolim"] += int(sale_quantity)

    elif sale_id == "Kit Autolim - Limpia Tapizados + Paño De Microfibra":
        total_full_sales["autolim"] += int(sale_quantity)
        total_full_sales["micro"]   += int(sale_quantity)

    elif sale_id == "Pack X 3 Limpia Tapizados Telas Cuero Butacas Techo Autolim":
        total_full_sales["autolim"] += 3 * int(sale_quantity)
        
    elif sale_id == "Kit Autolim X 2 Limpia Tapizados + 2 Paños De Microfibra":
        total_full_sales["autolim"] += 2 * int(sale_quantity)
        total_full_sales["micro"]   += 2 * int(sale_quantity)
    
    elif sale_id == "Limpiador Multiproposito + 2 Paños De Microfibra Autolim !":
        total_full_sales["autolim"] += int(sale_quantity)
        total_full_sales["micro"]   += 2 * int(sale_quantity)

    elif sale_id == "Kit Autolim - 2 Limpia Tapizados + 2 Paños De Microfibra Xl":
        total_full_sales["autolim"] += 2 * int(sale_quantity)
        total_full_sales["xl"]      += 2 * int(sale_quantity)

    elif sale_id == "Paño De Microfibra Autolim 37,5 X 37,5 Cm Limpieza Interior!":
        total_full_sales["micro"]   += int(sale_quantity)

    elif sale_id == "Paño De Microfibra Tamaño Xl Autolim - 61 X 77 Cm !!!":
        total_full_sales["xl"]   += int(sale_quantity)

    else: 
        no_match.append(full_sales[i])

print("")

print ("COMUN:")
print ("    Autolim:     " + str(total_common_sales["autolim"]))
print ("    microfibras: " + str(total_common_sales["micro"]))
print ("    XL:          " + str(total_common_sales["xl"]))
print ("    cepillos:    " + str(total_common_sales["cepillo"]))

print("")

print ("FULL:")
print ("    Autolim:     " + str(total_full_sales["autolim"]))
print ("    microfibras: " + str(total_full_sales["micro"]))
print ("    XL:          " + str(total_full_sales["xl"]))
print ("    cepillos:    " + str(total_full_sales["cepillo"]))

print("")

print ("TOTAL:")
print ("    Autolim:     " + str(total_full_sales["autolim"] + total_common_sales["autolim"]))
print ("    microfibras: " + str(total_full_sales["micro"] + total_common_sales["micro"]))
print ("    XL:          " + str(total_full_sales["xl"] + total_common_sales["xl"]))
print ("    cepillos:    " + str(total_full_sales["cepillo"] + total_common_sales["cepillo"]))

if len(no_match) > 1: # porque empieza en i = 1
    print("")
    print("------------------------------")
    print("ATENCION:")
    print(" NO MATCHEARON LOS SIGUIENTES:")
    print(no_match)

#print (full_sales)


