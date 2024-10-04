import pandas as pd
from datetime import datetime
import os
 
#si el elemento tiene un largo menor a 10 agregar 0s al principio
def agregar_ceros(elemento, largo, tipo):
   
    elemento = str(elemento)
   
    if tipo == "A":
        if len(elemento) <= largo:
            elemento = elemento + " "*(largo-len(elemento))
        else:
            elemento = elemento[:largo]
           
    elif tipo == "N":
        if len(elemento) <= largo:
            elemento = "0"*(largo-len(elemento)) + elemento
        else:
            elemento = elemento[:largo]
           
    return elemento
 
def crear_TR(fecha, beneficiario, contador):
   
    año = fecha[:2]
    mes = fecha[2:4]
    dia = fecha[4:]
   
    tr_code = año + mes + dia + str(contador) + beneficiario
    tr_seq_n = str(contador)
   
    return tr_code, tr_seq_n
 
def determinar_tipo_cuenta(cuenta):
   
    cuenta = str(cuenta).lower()

    if cuenta == "cuenta corriente" or cuenta == "cuenta vista":
        return "01"
    elif cuenta == "cuenta de ahorro":
        return "02"
    else:
        return "03"

def parsear_rut(rut):
    
    rut = str(rut).replace(".", "").replace("-", "").replace(" ", "")

    return rut
 
def parcear(ruta):
   
    df = pd.read_excel(ruta)
   
    cadenas = []
    total_pagar = 0
    contador = 0
   
    for _, row in df.iterrows():
       
        try:
            cadena = []
            
            fecha = row["Fecha de pago"]
            pais = row["Pais"]
            beneficiario = row["Beneficiario"]
            importe = str(row["Importe"]).split(",")[0].replace(".", "")
            descripcion = row["Concepto/Nº Factura"]
            sbif = row["Banco/Código Banco:"]
            cuenta = row["Cta.Banco:"]
            rut = parsear_rut(row["RUC/RUT/NIT"])
            pago = row["Forma de pago"]

            tipo_cuenta = determinar_tipo_cuenta(row["Tipo de Cuenta"])
            cuenta_pagadora = row["Número de cuenta pagadora"]
            divisa = row["Divisa (moneda paga)"]

            fecha = str(datetime.strptime(str(fecha), "%Y-%m-%d %H:%M:%S").strftime("%Y%m%d"))[2:]

            cuenta_pagadora = str(cuenta_pagadora).split(".", 1)[0]
            cuenta_pagadora = agregar_ceros(cuenta_pagadora, 10, "N")
            rut = agregar_ceros(rut, 10, "N")
            rut1 = agregar_ceros(rut, 20, "A")
            rut0 = agregar_ceros(rut, 16, "A")
            cuenta = agregar_ceros(cuenta, 35, "A")
            importe = agregar_ceros(importe, 15, "N")
            descripcion = agregar_ceros(descripcion, 35, "A")
            beneficiario = agregar_ceros(beneficiario, 80, "A")
            sbif = agregar_ceros(sbif, 3, "N")

            ad3 = agregar_ceros("00000001", 12, "A")

            tr_code,  tr_seq_n = crear_TR(fecha, beneficiario, contador)

            tr_code = agregar_ceros(tr_code, 15, "A")
            tr_seq_n = agregar_ceros(tr_seq_n, 8, "N")

            if divisa == "CLP":
            
                total_pagar += int(importe)
                contador += 1

                cadena.append("PAY") #OK
                cadena.append("152") #OK
                cadena.append(cuenta_pagadora) #OK
                cadena.append(fecha) #OK
                cadena.append("071") #OK
                cadena.append(tr_code) #OK
                cadena.append(tr_seq_n) #OK
                cadena.append("0000" + rut0) #OK
                cadena.append(divisa) #OK
                cadena.append(rut1) #OK
                cadena.append(importe) #OK
                cadena.append("      ") #OK
                cadena.append(descripcion) #OK
                cadena.append(" "*(35+35+35)) #OK
                cadena.append("01") #OK (seimpre 01?)
                cadena.append("01") #OK
                cadena.append(beneficiario) #OK
                cadena.append("SANTIAGO" + " "*27) #Direccion 1
                cadena.append(" "*35)  #OK
                cadena.append("SANTIAGO" + " "*7) #Direccion 2
                cadena.append("RM") #OK
                cadena.append(ad3) #OK
                cadena.append("0"*(16)) #OK
                cadena.append(sbif) #OK
                cadena.append("SANTIAGO") #OK
                cadena.append(cuenta) #OK
                cadena.append(tipo_cuenta) #OK
                cadena.append("SANTIAGO" + " "*22) #OK
                cadena.append("0"*(2 + 3)) #OK
                cadena.append(" "*14) #OK
                cadena.append("0"*(3)) #OK
                cadena.append(" "*19) #OK
                cadena.append("0"*(16)) #OK
                cadena.append(" "*(20 +15)) #OK
                cadena.append("0"*(10)) #OK
                cadena.append("01") #Preguntar Isabel
                cadena.append("001") #OK
                cadena.append(" "*50) #OK
                cadena.append("00122") #OK
                cadena.append(" "*50) #OK
                cadena.append("999999999999999") #OK
                cadena.append("2") #OK (primera vez deberia ser 2?)
                cadena.append("0"*11) #OK
                cadena.append(" "*(1 + 1)) #OK
                cadena.append("C") #OK
                cadena.append(" "*(15 + 238)) #OK
                        
                linea = ""
            
                for elemento in cadena:
                    linea += elemento
                
                cadenas.append(linea)
        except Exception as e:
            print("--------------------------------------------------------------------------------------------")
            print("Error en la linea: ", contador + 1, ":")
            print(e)
            print("--------------------------------------------------------------------------------------------")
    
    print("Lineas procesadas: ", contador)
    print("Total a pagar: $", total_pagar)  
    return cadenas, str(total_pagar)
 
def get_totales(cadenas, total_pagar):
   
   
    totales = ""
   
    total_lineas = agregar_ceros(len(cadenas), 15, "N")
    total_pagar = agregar_ceros(total_pagar, 15, "N")
   
    totales += "TRL" #OK
    totales += total_lineas #OK
    totales += total_pagar #OK
    totales += "0"*15 #OK
    totales += total_lineas #OK
    totales += " "*37 #OK
   
    return totales
 
#Escribe el archivo de texto
def write_file(cadenas, totales):
   
    with open("output.txt", "w") as file:
       
        for cadena in cadenas:
            file.write(cadena + "\n")
           
        file.write(totales)
       
    return
 
#lee todos los archivos de la carpeta
def read_all_files(ruta):
   
 
    files = os.listdir(ruta)
    for file in files:
        file_path = os.path.join(ruta, file)
       
       
        if os.path.isfile(file_path):
           
            cadenas, total_pagar = parcear(file_path)
            totales = get_totales(cadenas, total_pagar)
            write_file(cadenas, totales)
           
    return
ruta ="FICHEROS DE PAGO"
 
read_all_files(ruta)
 