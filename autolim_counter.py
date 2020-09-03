import xlrd
import os
import sys
import smtplib
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

current_directory = os.getcwd()

#print (current_directory)
#filePath = "/home/nahuel/Documents/autolim/ventas.xlsx"

filePath = current_directory + "/ventas.xlsx"
openFile = xlrd.open_workbook(filePath)
sheet = openFile.sheet_by_name("Ventas AR")

mes  = input('Ingrese el mes que desea analizar: ')

common_sales = []
full_sales = []
total_full_sales = {'autolim':0, 'micro':0, 'xl':0, 'cepillo':0 }
total_common_sales = {'autolim':0, 'micro':0, 'xl':0, 'cepillo':0 }
column_date = 1
column_deliver_ok = 2
column_shipping = 25
column_id = 14
column_quantity = 5
column_sale_id_ml = 0
column_buyer_name = 17
column_buyer_last_name = 18

# "set" of ids of only one Autolim
x1_autolim = ["Limpia Tu Tapizado Autolim Telas Cueros Alfombra Plasticos","Limpia Tu Tapizado  Autolim Telas Cueros Alfombra Plasticos","Limpiador Sillones Chenille Pana Tela Gamuza Cuero Autolim","Crema Limpia Tapizados Telas Cuero Butacas Detailing Autolim","Limpia Tapizados Telas Cuero Butacas Techo Detailing Autolim","Limpiador Tapizados Sillones Chenille Pana Tela Autolim",]
no_match = []
# print("Num de filas:", sheet.nrows ) # example to print number of rows

# iterate all sales, adding the ones for the month i want to their respective array (common or full)
for i in range(sheet.nrows ):

    cellValue = sheet.cell_value(i,column_date)
    if mes in str(cellValue): # okDate

        cellValue_shipping = sheet.cell_value(i,column_shipping)
        cellValue_id = sheet.cell_value(i,column_id)
        cellValue_quantity = sheet.cell_value(i,column_quantity)
        cellValue_deliver_ok = sheet.cell_value(i,column_deliver_ok)
        cellValue_sale_id_ml = sheet.cell_value(i,column_sale_id_ml)
        cellValue_buyer_name = sheet.cell_value(i,column_buyer_name)
        cellValue_buyer_last_name = sheet.cell_value(i,column_buyer_last_name)

        if not("Cancelada" in cellValue_deliver_ok):         # "Cancelada por el comprador"
            if "Full" in cellValue_shipping : # okFull
                full_sales.append((cellValue_id, cellValue_quantity, cellValue_sale_id_ml, cellValue_buyer_name, cellValue_buyer_last_name))
            else:
                common_sales.append((cellValue_id, cellValue_quantity, cellValue_sale_id_ml, cellValue_buyer_name, cellValue_buyer_last_name))





def calculateSales( sales, total_sales):
    for i in range(len(sales)):
        sale_id = sales[i][0]
        sale_quantity = sales[i][1]
        sale_id_ml = sales[i][2]
        sale_full_name = sales[i][3] + " " + sales[i][4]

        if sale_id in x1_autolim:
            total_sales["autolim"] += int(sale_quantity)

        elif sale_id == "Kit Autolim - Limpia Tapizados + Paño De Microfibra":
            total_sales["autolim"] += int(sale_quantity)
            total_sales["micro"]   += int(sale_quantity)
            #print(" 1 Autolim, 1 Paño x (" + sale_quantity + "), venta: #" + sale_id_ml + " " + sale_full_name )

        elif sale_id == "Pack X 3 Limpia Tapizados Telas Cuero Butacas Techo Autolim":
            total_sales["autolim"] += 3 * int(sale_quantity)
            
        elif sale_id == "Kit Autolim X 2 Limpia Tapizados + 2 Paños De Microfibra":
            total_sales["autolim"] += 2 * int(sale_quantity)
            total_sales["micro"]   += 2 * int(sale_quantity)
            #print(" 2 Autolim, 2 Paño x (" + sale_quantity + "), venta: #" + sale_id_ml + " " + sale_full_name )
        
        elif sale_id == "Limpiador Multiproposito + 2 Paños De Microfibra Autolim !":
            total_sales["autolim"] += int(sale_quantity)
            total_sales["micro"]   += 2 * int(sale_quantity)
            #print(" 1 Autolim, 2 Paño x (" + sale_quantity + "), venta: #" + sale_id_ml + " " + sale_full_name )

        elif sale_id == "Kit Autolim - 2 Limpia Tapizados + 2 Paños De Microfibra Xl":
            total_sales["autolim"] += 2 * int(sale_quantity)
            total_sales["xl"]      += 2 * int(sale_quantity)

        elif sale_id == "Paño De Microfibra Autolim 37,5 X 37,5 Cm Limpieza Interior!":
            total_sales["micro"]   +=  int(sale_quantity)
            #print(" 0 Autolim, 1 Paño x (" + sale_quantity + "), venta: #" + sale_id_ml + " " + sale_full_name )

        elif sale_id == "Paño De Microfibra Tamaño Xl Autolim - 61 X 77 Cm !!!":
            total_sales["xl"]   += int(sale_quantity)

        else: 
            no_match.append(sales[i])


def printAll():
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
    print("")
    print("")

    if len(no_match) > 1: # porque empieza en i = 1
        print("")
        print("------------------------------")
        print("ATENCION:")
        print(" NO MATCHEARON LOS SIGUIENTES:")
        print(no_match)
        print("")

def sendEmail(message):
    # create message object instance
    msg = MIMEMultipart()
    
    
    #message = "Thank you"
    recipients = ["EXAMPLE1@gmail.com","EXAMPLE2@gmail.com"]
    
    # setup the parameters of the message
    password = "EXAMPLEPASSWORD"
    msg['From'] = "EXAMPLE@gmail.com"
    msg['To'] = ", ".join(recipients)
    msg['Subject'] = "REPORTE VENTAS MERCADOLIBRE"
    
    # add in the message body
    msg.attach(MIMEText(message, 'plain'))
    
    #create server
    server = smtplib.SMTP('smtp.gmail.com: 587')
    
    server.starttls()
    
    # Login Credentials for sending the mail
    server.login(msg['From'], password)
    
    # send the message via the server.
    server.sendmail(msg['From'], msg['To'], msg.as_string())
    
    server.quit()


calculateSales(common_sales, total_common_sales)
calculateSales(full_sales, total_full_sales)
printAll()

desition_mail  = input('Queres enviarte este reporte via mail? (y/n): ')
print("")

if desition_mail == 'y' or desition_mail == 'Y':
    message = "FULL  : [Autolim: " +  str(total_full_sales["autolim"]) + ", Paños: "  +  str(total_full_sales["micro"]) + ", XL: "  +  str(total_full_sales["xl"]) + ", Cepillos: "  +  str(total_full_sales["cepillo"]) + " ]" + "\n" + "COMUN: [Autolim: " +  str(total_common_sales["autolim"]) + ", Paños: "  +  str(total_common_sales["micro"]) + ", XL: "  +  str(total_common_sales["xl"]) + ", Cepillos: "  +  str(total_common_sales["cepillo"]) + " ]"  
      

    sendEmail(message) 
    print ("Eviado correctamente :)")

elif desition_mail == 'n' or desition_mail == 'N':
    print("Ok, otra vez sera :(")

else:
    print ("Te dije que respondas Y o N, que parte no entendes?")




