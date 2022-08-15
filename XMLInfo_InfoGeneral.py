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

print('Leyendo cuantos XMLs hay...')


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

        #Relacionados
        try:
            Relacionados = myroot[[i for i, elem in enumerate(ElementNames) if 'Relacionado' in elem][0]]
            RelacionadosNames = [element.tag for element in Relacionados]
            RelacionadosUUIDs = [relacionado.attrib['UUID'] for relacionado in Relacionados]
        except:
            RelacionadosUUIDs = "-"

        # Datos de emisor
        Emisor=myroot[[i for i, elem in enumerate(ElementNames) if 'Emisor' in elem][0]]
        RFC=Emisor.attrib['Rfc']
        try:
            Nombre = Emisor.attrib['Nombre']
        except:
            Nombre = ""
        # DatosEmisor = [RFC,Nombre]

        # Importes
        SubTotal = myroot.attrib['SubTotal']
        Total = myroot.attrib['Total']
        try:
            Descuento = myroot.attrib['Descuento']
        except:
            Descuento = "0"
        # Importes = [SubTotal,Descuento,Total]

        # Salida
        # Salida = [UUID] + DatosEmisor + Importes + [path]

        Salida = {'UUID':UUID,
                'Relacionados':RelacionadosUUIDs,
                'RFC':RFC,
                'Razon social':Nombre,
                'Subtotal':SubTotal,
                'Descuento':Descuento,
                'Total':Total,
                'path':path}
    except:
        Salida = {'UUID':"-",
                'Relacionados':"-",
                'RFC':"-",
                'Razon social':"-",
                'Subtotal':"0",
                'Descuento':"0",
                'Total':"0",
                'path':path}

    try:
        return(pd.DataFrame.from_dict(Salida))
    except:
        return(pd.DataFrame(Salida, index=['0']))

#%%
print('Iniciando extracción...')

temp = []

for i in tqdm(range(len(xmlfiles))):
    temp.append(XMLInfo(xmldir[i],xmlfiles[i]))

XMLs=pd.concat(temp)

XMLs.to_excel('XMLs_InfoGeneral.xlsx')
