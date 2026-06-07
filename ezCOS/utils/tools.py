# -*- coding: utf-8 -*-
#
# tools.py -
#
# Authors: 
# Géry Casiez https://gery.casiez.net
#
# 2026
# 
# BSD License https://opensource.org/licenses/BSD-3-Clause
#
# Redistribution and use in source and binary forms, with or without
# modification, are permitted provided that the following conditions are met:
#
#  1. Redistributions of source code must retain the above copyright notice, this list of conditions
# and the following disclaimer.
#
# 2. Redistributions in binary form must reproduce the above copyright notice, this list of conditions
# and the following disclaimer in the documentation and/or other materials provided with the distribution.
# 
# 3. Neither the name of the copyright holder nor the names of its contributors may be used to endorse or
# promote products derived from this software without specific prior written permission.

# THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS" AND ANY EXPRESS OR IMPLIED WARRANTIES, 
# INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE ARE
# DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL,
# SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES;
# LOSS OF USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN 
# CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF THIS
# SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
#

import os
import json
from datetime import datetime
import pymupdf
import re

def getColumnInfo(ws) -> list:
    """
    Récupère les informations des colonnes d'une feuille du fichier Excel et les retourne sous forme de liste de dictionnaires.
    """
    list_with_values=[]
    for i, cell in enumerate(ws[1]):
        list_with_values.append({'column': cell.value, 'index': i})
    return list_with_values

def populateSheet(ws, infos, mapping):
    """
    Remplit la feuille Excel avec les données fournies dans le mapping.
     - ws : la feuille Excel à remplir
     - infos : les données à insérer dans la feuille
     - mapping : un dictionnaire qui associe les clés des données aux colonnes de la feuille Excel, avec éventuellement une transformation à appliquer aux données avant de les insérer.
    """

    columnsInfos = getColumnInfo(ws)
    row_index = 2
    for m in infos:
        for key, value in mapping.items():
            decomposed_key = key.split(':')
            data_value = m
            for k in decomposed_key:
                data_value = data_value.get(k, {})

            if data_value != {}:
                sheet_col_index = next((col['index'] for col in columnsInfos if col['column'] == value['column']), None)
                if sheet_col_index is not None:
                    if 'transformation' in value:
                        ws.cell(row=row_index, column=sheet_col_index + 1, value=value['transformation'](data_value))
                    else:
                        ws.cell(row=row_index, column=sheet_col_index + 1, value=data_value)
                else:
                    print(f"Colonne '{value['column']}' non trouvée dans le fichier Excel.")
        row_index += 1

def formatPrenom(prenom: str) -> str | None:
    """
    Formate le prénom en capitalisant la première lettre de chaque mot et en séparant les mots par des tirets.
    Par exemple, "jean-paul" devient "Jean-Paul" et "marie claire" devient "Marie Claire".
    """
    if prenom:
        stepOne = " ".join([p.capitalize() for p in prenom.split(" ")])
        return "-".join([p.capitalize() for p in stepOne.split("-")])
    return None

def formatLieu(lieu: str) -> str | None:
    """
    Formate le lieu en capitalisant la première lettre de chaque mot, sauf pour les mots "DE" qui sont mis en majuscules."
    """
    if lieu:
        return " ".join([l.capitalize() if l not in ["DE"] else l.lower() for l in lieu.split(" ")])
    return None

def computeAge(d) -> float:
    """
    Calcule l'âge à partir d'une date de naissance donnée au format "dd/mm/yyyy". Si la date est "nan", retourne 0.
    """
    if str(d) != "nan":
        a = datetime.now()
        b = datetime.strptime(d,"%d/%m/%Y")
        c = a - b
        return c.days/365.0
    else:
        return 0

def loadJson(filename):
    """
    Charge les données d'un fichier JSON et les retourne sous forme de dictionnaire.
    """
    with open(filename, 'r', encoding='utf8') as f:
        return json.load(f)

def saveJson(data, filename):
    """
    Sauvegarde les données dans un fichier JSON avec une indentation de 4 espaces et en préservant les caractères non ASCII.
    """
    with open(filename, 'w', encoding='utf8') as f:
        json.dump(data, f, indent=4, ensure_ascii=False)

def createDir(path):
    """Crée un répertoire à l'emplacement spécifié s'il n'existe pas déjà.
    Args:        
        path (str): Le chemin du répertoire à créer.
    """
    if (not(os.path.exists(path))):
        os.makedirs(path)  

def getTextFromPDF(file):
    """
    Extrait le texte d'un fichier PDF en utilisant la bibliothèque pymupdf.
    Args:
        file (str): Le chemin du fichier PDF.
    Returns:
        str: Le texte extrait du PDF.
    """
    pdf_document = pymupdf.open(file)
    text = ""

    for page_num in range(pdf_document.page_count):
        page = pdf_document.load_page(page_num)
        text += page.get_text()
    return text

def extractInfo(info: str, txt: str, stopWord=None) -> str:
    """
    Extrait une information spécifique d'un texte en utilisant une expression régulière.
    Args:
    - info (str): L'information à extraire, utilisée comme point de départ dans le texte.
    - txt (str): Le texte à partir duquel extraire l'information.
    - stopWord (str, optional): Un mot qui indique la fin de l'information à extraire. Si None, l'extraction se fait jusqu'à la fin de la ligne.
    Returns:
    str: L'information extraite du texte, ou une chaîne vide si l'information n'est pas trouvée.
    """
    res = ""

    if stopWord:
        pattern = f"{info} : (.*?){stopWord}"
        match = re.search(pattern, txt, re.DOTALL)
    else:
        pattern = f"^{info} : (.*)$"
        match = re.search(pattern, txt, re.MULTILINE)

    if match:
        res = match.group(1).strip()
        if res is not None:
            res = res.replace("\n", "")

    return res

def extractInfoRapport(info, txt, stopWord=None):
    res = ""

    if stopWord:
        pattern = f"{info}(.*?){stopWord}"
        match = re.search(pattern, txt, re.DOTALL)
    else:
        pattern = f"{info}(.*?)$"
        match = re.search(pattern, txt, re.MULTILINE)

    if match:
        res = match.group(1).strip()
        if res is not None:
            res = res.replace("\n", "")

    return res


def formattxt(txt, prefix = " - "):
	if (str(txt) == "") or (str(txt) == "nan"):
		return ""
	else:
		return "%s%s"%(prefix, txt)
    
def cleantext(txt, convBullets=False):
    if str(txt) != "nan":
        res = str(txt).replace("\"","").replace("   ","").replace("_x000D_","").replace("¿","-")
        if convBullets:
            return res.replace(" -","\n-").replace(",-",",\n-")
        else:
            return res
    else:
        return ""
    
def thunderbird(thunderbirdbin, recipientName, recipientAddress, thesubject, thecontent, cc='' , attachment=''):
    cmd = "%s -compose \"to='%s',cc='%s',subject='%s',format=1,body='%s',attachment='%s'\""%(thunderbirdbin, recipientAddress, cc, thesubject, thecontent, attachment)
    # print(cmd)
    os.system(cmd)

def trAvis(avis):
    if avis == "TF":
        return "Très favorable"
    elif avis == "F":
        return "Favorable"
    elif avis == "R":
        return "Réservé"
    elif avis == "D":
        return "Défavorable"
    else:
        return avis