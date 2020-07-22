import requests

#http://portalbissrs.xm.com.co/dmnd/Paginas/Historicos/Historicos.aspx

file = 'Demana_Comercial_SEM2_2020.xlsx'
link = 'http://portalbissrs.xm.com.co/dmnd/Histricos/Demanda_Comercial_Por_Comercializador_SEME2_2020.xlsx'

print("")
resp = requests.get(link)
output = open(file, 'wb')
output.write(resp.content)
output.close()
print("Archivo descargado:",file)
