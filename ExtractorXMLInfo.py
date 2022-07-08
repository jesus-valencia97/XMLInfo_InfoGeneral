# -*- coding: utf-8 -*-
"""
Created on Thu Mar 17 11:24:53 2022

@author: -
"""

#%%
import pandas as pd
import numpy as np
import os
import xml.etree.ElementTree as ET
import win32api
from tqdm import tqdm

#%%
os.chdir(os.path.abspath(os.path.dirname(__file__)))
print(os.getcwd())

#%%
xmlfiles=list()
xmldir=list()
for root, dirs, files in os.walk(os.getcwd()):
    for file in files:
        if file.lower().endswith(".xml"):
            #xmlpaths.append(os.path.join(root, file))
            xmlfiles.append(file)
            xmldir.append(root)
            #print(os.path.join(root, file))

#%%
def ExtractUUID(direc,file):
    path= direc + '\\' + file
    shortpath= win32api.GetShortPathName(direc) + '\\' + file
    try:
        mytree = ET.parse(shortpath)
        myroot = mytree.getroot()
        ElementNames=[element.tag for element in myroot]

        #UUID
        Complemento = myroot[[i for i, elem in enumerate(ElementNames) if 'Complemento' in elem][0]]
        ComplementNames = [element.tag for element in Complemento]
        TimbreDigital = Complemento[[i for i, elem in enumerate(ComplementNames) if 'Timbre' in elem][0]]
        UUID = TimbreDigital.attrib['UUID']
    except:
        UUID = ''

    return UUID, direc, file

#%%
print('Extrayendo valores únicos...')

temp = []

for i in tqdm(range(len(xmlfiles))):
    temp.append(ExtractUUID(xmldir[i],xmlfiles[i]))

UUIDs=pd.DataFrame(temp,columns = ['UUID','direc','file'])

#%%
UUIDs = UUIDs.drop_duplicates(subset=['UUID'])
xmlfiles=list(UUIDs['file'])
xmldir=list(UUIDs['direc'])


#%%
def XMLInfo(direc,file):
    path= direc + '\\' + file
    shortpath= win32api.GetShortPathName(direc) + '\\' + file
    try:
        mytree = ET.parse(shortpath)
        myroot = mytree.getroot()
        ElementNames=[element.tag for element in myroot]

        #UUID
        Complemento = myroot[[i for i, elem in enumerate(ElementNames) if 'Complemento' in elem][0]]
        ComplementNames = [element.tag for element in Complemento]
        TimbreDigital = Complemento[[i for i, elem in enumerate(ComplementNames) if 'Timbre' in elem][0]]
        UUID = TimbreDigital.attrib['UUID']

        # Datos de emisor
        Emisor=myroot[[i for i, elem in enumerate(ElementNames) if 'Emisor' in elem][0]]
        RFC=Emisor.attrib['Rfc']
        try:
            Nombre = Emisor.attrib['Nombre']
        except:
            Nombre = "-"
        DatosEmisor = [RFC,Nombre]

        # Importes

        ## Conceptos

        Conceptos = myroot[[i for i, elem in enumerate(ElementNames) if 'Conceptos' in elem][0]]

        ConceptosList=[]

        ### X Concepto
        for concepto in Conceptos:
            
            cNames = [element.tag for element in concepto]
            # Impuestos
            
            try:
                cImpuestos = concepto[[i for i, elem in enumerate(cNames) if 'Impuestos' in elem][0]]
                
                cImpuestosNames = [element.tag for element in cImpuestos]
            
                Traslados = {}
            
                # Traslados

                cTraslados = cImpuestos[[i for i, elem in enumerate(cImpuestosNames) if 'Traslados' in elem][0]]
                
                for i,traslado in enumerate(cTraslados):
                    Impuesto = traslado.attrib['Impuesto']
                    Base = traslado.attrib['Base']
                    TasaOCuota = traslado.attrib['TasaOCuota']
                    Importe = traslado.attrib['Importe']

                    # Diccionario
                    
                    t_Traslados = {}
                    t_Traslados['Impuesto']=Impuesto
                    t_Traslados['Base']=float(Base)
                    t_Traslados['TasaOCuota']=float(TasaOCuota)
                    t_Traslados['Importe']=float(Importe)
                    
                    Traslados[Impuesto]=t_Traslados
                    
                Traslados=pd.DataFrame(Traslados).T
                Traslados = Traslados[['Impuesto','Base','TasaOCuota','Importe']]
                Traslados.columns = pd.MultiIndex.from_arrays([['Traslados']*4,['Impuesto','Base','TasaOCuota','Importe']]) 
                
            except:
            
                Traslados = pd.DataFrame(columns=pd.MultiIndex.from_arrays([['Traslados']*4,['Impuesto','Base','TasaOCuota','Importe']]))   

            
            # Retenciones
            
            
            try:
                Retenciones = {}
                
                cRetenciones = cImpuestos[[i for i, elem in enumerate(cImpuestosNames) if 'Retenciones' in elem][0]]
                                        
                for i,retencion in enumerate(cRetenciones):
                    Impuesto = retencion.attrib['Impuesto']
                    Base = retencion.attrib['Base']
                    TasaOCuota = retencion.attrib['TasaOCuota']
                    Importe = retencion.attrib['Importe']

                    # Diccionario
                    

                    t_Retenciones = {}
                    t_Retenciones['Impuesto']=Impuesto
                    t_Retenciones['Base']=float(Base)
                    t_Retenciones['TasaOCuota']=float(TasaOCuota)
                    t_Retenciones['Importe']=float(Importe)
                    
                    Retenciones[Impuesto]=t_Retenciones
                    
                Retenciones=pd.DataFrame(Retenciones).T
                Retenciones = Retenciones[['Impuesto','Base','TasaOCuota','Importe']]
                Retenciones.columns = pd.MultiIndex.from_arrays([['Retenciones']*4,['Impuesto','Base','TasaOCuota','Importe']]) 
                
            except:
            
                Retenciones = pd.DataFrame(columns=pd.MultiIndex.from_arrays([['Retenciones']*4,['Impuesto','Base','TasaOCuota','Importe']]))
                # Retenciones=pd.DataFrame(Retenciones).T
                # Retenciones = Retenciones[['Impuesto','Base','TasaOCuota','Importe']]
                # Retenciones.columns = pd.MultiIndex.from_arrays([['Retenciones']*4,['Impuesto','Base','TasaOCuota','Importe']])           
         
            Importe_concepto = float(concepto.attrib['Importe'])
            try:
                Descuento_concepto = float(concepto.attrib['Descuento'])
            except:
                Descuento_concepto = 0   
            
            Descripcion_concepto = concepto.attrib['Descripcion']
                
                
            if Traslados.shape[0] + Retenciones.shape[0]==0:
                t_Concepto = pd.DataFrame([[Descripcion_concepto,Importe_concepto,Descuento_concepto]],columns=['Descripción','Importe','Descuento'])
                t_Concepto = pd.concat([t_Concepto,Traslados.join(Retenciones,how='outer')])
                
            else:
                t_Concepto = Traslados.join(Retenciones,how='outer')
                t_Concepto.insert(0, 'Descuento', Descuento_concepto)
                t_Concepto.insert(0, 'Importe', Importe_concepto)
                t_Concepto.insert(0, 'Descripción', Descripcion_concepto)
            
            t_Concepto.columns = ['Descripción','Importe','Descuento',
                'T_Impuesto','T_Base','T_TasaOCuota','T_Importe',
                'R_Impuesto','R_Base','R_TasaOCuota','R_Importe']
            
            ConceptosList.append(t_Concepto)
            
        CFDI = pd.concat(ConceptosList)      
            
        ## Totales


        SubTotal = float(myroot.attrib['SubTotal'])
        try:
            Descuento = float(myroot.attrib['Descuento'])
        except:
            Descuento = 0
        Total = float(myroot.attrib['Total'])

        # Total impuestos

        try:
            Impuestos = myroot[[i for i, elem in enumerate(ElementNames) if 'Impuestos' in elem][0]]
            
            # Traslados
            try: 
                tot_traslados = float(Impuestos.attrib['TotalImpuestosTrasladados'])
            except:
                tot_traslados = np.nan
                
            # Retenciones
            try: 
                tot_retenciones = float(Impuestos.attrib['TotalImpuestosRetenidos'])
            except:
                tot_retenciones = np.nan
                
            t_total=pd.DataFrame(['Total',SubTotal,Descuento,np.nan,np.nan,np.nan,tot_traslados,np.nan,np.nan,np.nan,tot_retenciones]).T

        except:
            t_total=pd.DataFrame(['Total',SubTotal,Descuento,np.nan,np.nan,np.nan,np.nan,np.nan,np.nan,np.nan,np.nan]).T

        t_total.columns = CFDI.columns

        CFDI = pd.concat([CFDI,t_total])

        CFDI['Total_CFDI'] = Total

        CFDI.insert(0, 'Razón social', Nombre)
        CFDI.insert(0, 'RFC', RFC)
        CFDI.insert(0, 'UUID', UUID)

        CFDI['path'] = path
        
    except:
        # tuples = [(          'UUID',           ''),
        #             (           'RFC',           ''),
        #             (  'Razón social',           ''),
        #             (   'Descripción',           ''),
        #             (       'Importe',           ''),
        #             (     'Descuento',           ''),
        #             (     'Traslados',   'Impuesto'),
        #             (     'Traslados',       'Base'),
        #             (     'Traslados', 'TasaOCuota'),
        #             (     'Traslados',    'Importe'),
        #             (   'Retenciones',   'Impuesto'),
        #             (   'Retenciones',       'Base'),
        #             (   'Retenciones', 'TasaOCuota'),
        #             (   'Retenciones',    'Importe'),
        #             (    'Total_CFDI',           ''),
        #             (          'path',           '')]
        tuples = ['UUID','RFC','Razón social','Descripción','Importe','Descuento',
               'T_Impuesto','T_Base','T_TasaOCuota','T_Importe',
               'R_Impuesto','R_Base','R_TasaOCuota','R_Importe',
               'Total_CFDI',
               'path']
        # columnsnames = pd.MultiIndex.from_tuples(tuples)
        
        CFDI = pd.DataFrame(columns = tuples)
    
    return(CFDI)



#%%
print('Iniciando extracción...')

temp = []

for i in tqdm(range(len(xmlfiles))):
    temp.append(XMLInfo(xmldir[i],xmlfiles[i]))

XMLs=pd.concat(temp)

# XLSXcolumns = ['UUID','RFC','Razón social','Descripción','Importe','Descuento',
#                'T_Impuesto','T_Base','T_TasaOCuota','T_Importe',
#                'R_Impuesto','R_Base','R_TasaOCuota','R_Importe',
#                'Total_CFDI',
#                'path']
              
# XMLs.columns = XLSXcolumns

XMLs.insert(6,'IMPORTE',XMLs['Importe']-XMLs['Descuento'])
#%%
XMLs_total=XMLs.loc[XMLs['Descripción']=='Total'].reset_index(drop=True)
# XMLs_total = XMLs_total.drop_duplicates(subset=set(XLSXcolumns)-set(['path']))
XMLs_detail=XMLs.loc[XMLs['Descripción']!='Total'].reset_index(drop=True)
# XMLs_detail = XMLs_detail.drop_duplicates(subset=set(XLSXcolumns)-set(['path']))

#%%
with pd.ExcelWriter('XMLs_wImpuestos.xlsx') as writer:
    XMLs_detail.to_excel(writer, sheet_name='Detalle',freeze_panes=(1,0))
    XMLs_total.to_excel(writer, sheet_name='Totales',freeze_panes=(1,0))
