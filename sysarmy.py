# -*- coding: utf-8 -*-
"""
Created on Tue Sep 12 07:56:45 2023

@author: sosafmatias
Analisis de base de datos Sysarmy 2023-2. Objetivos:
    1- Participacion 
        por provincia
        Puesto de trabajo
        Sexo
        Contrato
    
    2- Agrupaciones
        Participantes por provincia con grafico
                    por tipo de contrato
                    por genero con grafico
                    cantidad por puesto y provincia
                    promedio de salario por provincia y puesto

"""
import pandas as pd

path="D:/Data Science/Practicas/Python/"

#importo mi dataset, tengo los títulos de columnas en la fila 8 de Excel
#quiero trabajarlo con el engine de openpyxl
sueldos_df=pd.read_excel(
    path+"Dataset Sysarmy.xlsx",
    header=[7],
    engine="openpyxl"
    )

#creo la variable donde voy a exportar el archivo nuevo
writer=pd.ExcelWriter(path+"Analisis Sysarmy.xlsx")

#filtro para quedarme las columnas que me interesan
#renombro usando method chaining por nombres intuitivos
analisis_df=sueldos_df[[
    "Dónde estás trabajando", "Dedicación", "Trabajo de", "Tiempo en el puesto actual",
    "Último salario mensual  o retiro BRUTO (en tu moneda local)", "Me identifico (género)" 
    ]].rename(columns={
        "Dónde estás trabajando":"Provincia",
        "Dedicación":"Contrato",
        "Trabajo de":"Puesto de trabajo",
        "Tiempo en el puesto actual":"Antiguedad en el puesto",
        "Último salario mensual  o retiro BRUTO (en tu moneda local)":"Salario bruto",
        "Me identifico (género)":"Genero"
        })

#Encontré en el dataset un encuestado que NO trabaja en IT, si no que es su pretensión.
#Entiendo que desvirtúa la encuesta al no ser un salario del rubro pedido, lo busco y elimino
indice_eliminar = analisis_df[analisis_df["Puesto de trabajo"
                ]=="soy de finanzas estoy buscado trabajo de Front end o QA tester, sql y react, actualmente estudio programacion por ende quiero cambiar de área al mundo IT que tanto me apasiona. He realizado proyectos de emulación, bajo la metodología SCRUM. Gracias."].index[0]
analisis_df.drop(analisis_df.index[indice_eliminar],inplace=True)

#Necesido formatear todos los textos de puesto de trabajo para tener consistencia en las respuestas
analisis_df["Puesto de trabajo"]=analisis_df["Puesto de trabajo"].str.title()

     

#Este dataset tiene entrada de texto libre y tengo 639 Géneros únicos
#Voy a reemplazar todo lo que no sea Varón Cis o Mujer Cis por "Otro"
#Elijo conservar estos 2 porque son mayoría con diferencia
analisis_df.loc[
    (analisis_df["Genero"]!="Varón Cis") & (analisis_df["Genero"]!="Mujer Cis"),
    ["Genero"]
    ]="Otro"

#Reemplazo salarios menores a 100000 por el promedio ya que algunos expresaron centenas como miles
#ej. "200" en lugar de "200000"
analisis_df.loc[
    (analisis_df["Salario bruto"]<100000),["Salario bruto"]
    ]=analisis_df["Salario bruto"].mean()


#Cantidad de participantes por Provincia
participacionxprovincia=analisis_df[
    ["Provincia","Puesto de trabajo"]
    ].groupby(
        ["Provincia"],
        sort=True
        ).count().rename(
                    columns={
                    "Puesto de trabajo":"Cantidad"
                    }
                    ).to_excel(
                    writer, 
                    sheet_name='Provincia',
                    engine='xlsxwriter'
                    )

            
##Cantidad de participantes por Tipo de contrato
participacionxcontrato=analisis_df[
    ["Contrato","Puesto de trabajo"]
    ].groupby(
        ["Contrato"],
        sort=True
        ).count().rename(
                    columns={
                    "Puesto de trabajo":"Cantidad"
                    }
                    ).to_excel(
                    writer, 
                    sheet_name='Tipo de contrato',
                    engine='xlsxwriter'
                    )
       
            
#Cantidad de participantes por Genero
participacionxgenero=analisis_df[
    ["Genero","Puesto de trabajo"]
    ].groupby(
        ["Genero"],
        ).count().rename(
                    columns={
                    "Puesto de trabajo":"Cantidad"
                    }
                    ).to_excel(
                    writer, 
                    sheet_name='Genero',
                    engine='xlsxwriter'
                    )


#Cantidad de participantes por Puesto de trabajo en cada provincia    
participacionxpuesto=analisis_df[
    ["Contrato","Puesto de trabajo","Provincia"]
    ].groupby(
        ["Puesto de trabajo","Provincia"]
        )["Puesto de trabajo"].count(
            ).to_excel(
                    writer, 
                    sheet_name='Puesto de trabajo',
                    engine='xlsxwriter'
                    )

#Salario bruto promedio por provincia, para cada puesto de trabajo
sueldoxprovincia=analisis_df[
    ["Salario bruto","Puesto de trabajo","Provincia"]
    ].groupby(
        ["Provincia","Puesto de trabajo"]
        )["Salario bruto"].mean(
            ).to_excel(
                    writer, 
                    sheet_name='Promedio salarial por puesto',
                    engine='xlsxwriter'
                    )


#Creo las pestañas en mi archivo de excel para tener cada reporte por separado
workbook  = writer.book
worksheet_participacionxprovincia = writer.sheets['Provincia']
worksheet_participacionxcontrato = writer.sheets['Tipo de contrato']
worksheet_participacionxgenero = writer.sheets['Genero']
worksheet_participacionxpuesto = writer.sheets['Puesto de trabajo']
worksheet_promedioxpuesto = writer.sheets['Promedio salarial por puesto']

#Formatos de celdas
bold = workbook.add_format({
    'bold': True,
    
    })

center = workbook.add_format({
    'align':'center',
    
    })

centerbold = workbook.add_format({
    'bold': True,
    'align':'center'
    })

left = workbook.add_format({
    'align':'left',
    
    })


##formatear decimales
float2decimals = workbook.add_format(
    {
     'num_format': '$#,##0.00'
    }
    )


#grafico de torta
chartpie = workbook.add_chart(
    {
     'type': 'pie',    
    }
    )


#tamaño en pixels del grafico dentro de la planilla
chartpie.set_size(
    {
     'width': 720,
     'height': 576
     }
    )


#Le doy al gráfico las series de datos necesarias
chartpie.add_series(
    {
     'name':'Participantes por provincia',
     'categories': '=Provincia!$A$2:$A$25',
    'values':     '=Provincia!$B$2:$B$25',
     }
    )

#grafico de columnas
chartcol = workbook.add_chart(
    {
     'type': 'column',    
    }
    )

chartcol.set_size(
    {
     'width': 500,
     'height': 350
     }
    )

chartcol.add_series(
    {
     'name':'Distribucion por Genero',
     'categories': '=Genero!$A$2:$A$4',
    'values':     '=Genero!$B$2:$B$4'
     }
    )

#Desactivo la leyenda del gráfico que es redundante con el título y hay un solo dato: Cantidad
chartcol.set_legend(
    {
     'none': True
     }
    )

#inserto el grafico propiamente con su extremo izquierdo en la coordenada y hoja deseada
worksheet_participacionxprovincia.insert_chart('D2', chartpie)
worksheet_participacionxgenero.insert_chart('D2', chartcol)

#Cambios de ancho de columnas, alineaciones y titulos
worksheet_participacionxprovincia.set_column("A:A",32)
worksheet_participacionxprovincia.set_column("B:B", 9, center)

worksheet_participacionxgenero.set_column('A:B', 9, center)

worksheet_participacionxcontrato.set_column('A:B', 9, center)

worksheet_participacionxpuesto.set_column("A:A",79)
worksheet_participacionxpuesto.set_column("B:B",32)
worksheet_participacionxpuesto.write("C1", "Cantidad", centerbold)
worksheet_participacionxpuesto.set_column("C:C", 16, center)

worksheet_promedioxpuesto.set_column("A:A",32)
worksheet_promedioxpuesto.set_column("B:B",79)
#Seteo de cantidad de decimales para sueldo bruto
worksheet_promedioxpuesto.set_column("C:C", 12, float2decimals)

#Guardo mi archivo en la ruta establecida previamente tanto en Excel como en DB para SQLite
writer.save()


