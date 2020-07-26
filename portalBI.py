import requests
import pandas as pd
import os
from bokeh.plotting import figure, output_file, show
from bokeh.layouts import row
import bokeh.palettes as pl
#http://portalbissrs.xm.com.co/dmnd/Paginas/Historicos/Historicos.aspx

def download_file(file, path_link):
    """
    función que recibe un link para descargar un archivo
    """
    link =  path_link + file
    resp = requests.get(link)
    output = open(file, 'wb')
    output.write(resp.content)
    output.close()    
    print("Archivo descargado:",file)

    
def read_files(file):
    """
    Función para cargar la informción de los archivos
    como DataFrame
    """
    rute = os.getcwd() + os.sep + file
    file_ = pd.ExcelFile(rute)
    df =  file_.parse(file_.sheet_names[0])
    df = df.rename(columns=df.iloc[1]).drop([0,1],axis=0).reset_index(drop=True)
    df = df.fillna(value=0)
    headers = list(df.columns)
    for k in range(len(headers)):
        if k >= 3 and headers[k] != 'Version':
            headers[k] = 'h'+str(headers[k])
    df.columns = headers

    regulado = df[df['Mercado'] == 'REGULADO'].groupby(['Fecha','Mercado'])\
    ['Mercado','h0', 'h1.0', 'h2.0', 'h3.0', 
     'h4.0', 'h5.0', 'h6.0', 'h7.0', 'h8.0', 'h9.0', 'h10.0', 'h11.0', 'h12.0', 
     'h13.0', 'h14.0', 'h15.0', 'h16.0', 'h17.0', 'h18.0',
     'h19.0', 'h20.0', 'h21.0', 'h22.0', 'h23.0', 'Version'].sum()

    no_regulado = df[df['Mercado'] == 'NO REGULADO'].groupby(['Fecha','Mercado'])\
   ['Mercado','h0', 'h1.0', 'h2.0', 'h3.0', 
     'h4.0', 'h5.0', 'h6.0', 'h7.0', 'h8.0', 'h9.0', 'h10.0', 'h11.0', 'h12.0', 
     'h13.0', 'h14.0', 'h15.0', 'h16.0', 'h17.0', 'h18.0',
     'h19.0', 'h20.0', 'h21.0', 'h22.0', 'h23.0', 'Version'].sum()

    consumo = df[df['Mercado'] == 'CONSUMO'].groupby(['Fecha','Mercado'])\
    ['Mercado','h0', 'h1.0', 'h2.0', 'h3.0', 
     'h4.0', 'h5.0', 'h6.0', 'h7.0', 'h8.0', 'h9.0', 'h10.0', 'h11.0', 'h12.0', 
     'h13.0', 'h14.0', 'h15.0', 'h16.0', 'h17.0', 'h18.0',
     'h19.0', 'h20.0', 'h21.0', 'h22.0', 'h23.0', 'Version'].sum()
    df_final = pd.concat([regulado, no_regulado, consumo]).reset_index()
    
    print('DataFrame Cargado')
    
    return df_final


def to_plot_all(list_fecha):
    """
    Filtra la información que sera gráficada
    """

    def data_to_plot(fecha, tipo):    

        data = dfs_[dfs_['Mercado'] == tipo]
        data_fecha = data[data['Fecha'] == fecha].T
        data_fecha = data_fecha.reset_index(drop=True)\
                        .drop([0,1],axis=0).reset_index(drop=True)

        data_fecha.columns = [tipo +'_'+fecha]
        if tipo != 'NO REGULADO':
            data_fecha['Hora'] = [i for i in range(0,24)]   

        return data_fecha
    
    to_plot = []
    
    for fecha in list_fecha:
        regulada_fecha = data_to_plot(fecha,'REGULADO')
        no_regulada_fecha = data_to_plot(fecha,'NO REGULADO')

        to_plot_1 = pd.concat([regulada_fecha, no_regulada_fecha], axis = 1)
        to_plot.append(to_plot_1)
        
    return to_plot   


def graficar(to_plot):
    """
    función que gráfica la demanda para una lista de días
    """

    #from bokeh.io import output_notebook, vplot
    if len(to_plot)>2:
        colors = pl.all_palettes['Viridis'][len(to_plot)]
    else: 
        colors = ['red','blue']
    
    output_file('graficado_simple.html')
    fig = figure(plot_width=800, plot_height=600)
    fig2 = figure(plot_width=800, plot_height=600)

    fig.title.text = "Demanda del SIN " + fecha_1 +" - "+fecha_2 + " REGULADA"
    fig.title.align = "center"
    fig.title.text_color = "black"
    fig.title.text_font_size = "25px"
    fig.title.background_fill_color = "#aaaaee"
    fig.xaxis.axis_label = 'HORA'
    fig.yaxis.axis_label = 'DEMANDA/1000'

    fig2.title.text = "Demanda del SIN " + fecha_1 +" - "+fecha_2 + " NO REGULADO"
    fig2.title.align = "center"
    fig2.title.text_color = "black"
    fig2.title.text_font_size = "25px"
    fig2.title.background_fill_color = "#aaaaee"
    fig2.xaxis.axis_label = 'HORA'
    fig2.yaxis.axis_label = 'DEMANDA/1000'
    
    col_counter = 0
    for plot in to_plot:        
        
        fig.circle(plot['Hora'],plot[list(plot.columns)[0]]/1000, line_width=2,
                 legend_label=list(plot.columns)[0],color = colors[col_counter])
        fig.line(plot['Hora'],plot[list(plot.columns)[0]]/1000, line_width=2,
                 legend_label=list(plot.columns)[0],color = colors[col_counter])

        fig2.circle(plot['Hora'],plot[list(plot.columns)[2]]/1000, line_width=2,
                 legend_label=list(plot.columns)[2],color = colors[col_counter])
        fig2.line(plot['Hora'],plot[list(plot.columns)[2]]/1000, line_width=2,
                 legend_label=list(plot.columns)[2],color = colors[col_counter])
        col_counter += 1
        
    show(row(fig,fig2))  
    
  
    
if __name__ == '__main__':
    
    #Lista de archivos que quiero descargar para la carpeta demanda
    
    #el path e spor si el archivo es de otra carpeta de XM
    path_link = 'http://portalbissrs.xm.com.co/dmnd/'\
        'Histricos/'
    
    files = ['Demanda_Comercial_Por_Comercializador_SEME1_2019.xlsx',
            'Demanda_Comercial_Por_Comercializador_SEME2_2019.xlsx',
            'Demanda_Comercial_Por_Comercializador_SEME1_2020.xlsx',
            'Demanda_Comercial_Por_Comercializador_SEME2_2020.xlsx']
    
    #Lista para guardar los DataFrame
    dfs = []
    #Leo a información de cada Excel
    for file in files:
        if os.path.exists(file):
            pass
        else:
            download_file(file,path_link)
        
        dfs.append(read_files(file))
    
    #Concateno toda la información
    dfs_ = pd.concat(dfs).reset_index(drop=True)
    #Elijo las fechas que quiero gráficar para comparar
    to_plot = to_plot_all(['2019-05-01','2019-04-01','2020-05-01','2020-04-01'])
    #Grafico las fechas
    graficar(to_plot)
    #Se guarda la información graficada en un Excel
    with pd.ExcelWriter('Demanda_output.xlsx') as writer:  

        (pd.concat(to_plot,axis=1)).to_excel(writer, sheet_name='Demanda')

    os.startfile('Demanda_output.xlsx')
    