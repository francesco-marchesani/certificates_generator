# coding=utf-8
from __future__ import print_function
from mailmerge import MailMerge
from datetime import date

"""
- Online reference (for Python + Word): https://pbpython.com/python-word-template.html

- Files needed: certificates_generator.py (python script) + POLIMI.docx (template document), they must be in the same folder (e.g., Desktop)
 
- Following dependendces have to be installed on Linux, MacOS and Windows: 
conda install lxml
pip install docx-mailmerge

"""

#Collecting inputs
print('\n\nWelcome to GSOM, this is the unofficial declaration of attendance generator coded by Francesco Marchesani \n\n ** Please write with the first capital letter, then lower sentence!\n\n')
name = raw_input('Write your name (e.g., Mario)\n')
surname = raw_input('Write your surname (e.g., Rossi)\n')
city = raw_input('Write the city where you were born (e.g., Milano)\n')
day_of_birth = raw_input('Write the day where you were born (e.g., 10)\n')
month_of_birth = raw_input('Write the month where you were born (e.g., 03)\n')
year_of_birth = raw_input('Write the year where you were born (e.g., 1990)\n')
country_of_birth = raw_input('Write the country where you were born (e.g., Italy)\n')


#Generating Innovation Strategy certificate
template = "POLIMI.docx"
document = MailMerge(template)
print(document.get_merge_fields())

document.merge(
    #Inserire la città di nascita
    CITY=city, 
    #Inserire il proprio cognome
    SURNAME=surname, 
    #Inserire il proprio nome
    NAME=name, 
    #Inserire il mese di nascita
    MM=month_of_birth,
    #Inserire la nazione di nascita
    COUNTRY=country_of_birth, 
    #Inserire il giorno di nascita
    DD=day_of_birth, 
    #Inserire l'anno di nascita
    YYYY=year_of_birth,
    #Inserire il nome del corso 
    COURSENAME='Innovation Strategy', 
    #Inserire il mese del corso
    ML='January', 
    #Inserire il mese di rilascio
    MFINAL='January', 
    #Inserire il secondo giorno di lezioni
    D2='14th', 
    #Inserire il giorno di rilascio dell'attestato
    DFINAL='16th', 
    #Inserire il primo giorno di lezioni 
    D1='13th')

document.write('innovation_strategy.docx')

#Generating Business Strategy & Project Management Certificate
template = "POLIMI.docx"
document = MailMerge(template)
print(document.get_merge_fields())

document.merge(
    #Inserire la città di nascita
    CITY=city, 
    #Inserire il proprio cognome
    SURNAME=surname, 
    #Inserire il proprio nome
    NAME=name, 
    #Inserire il mese di nascita
    MM=month_of_birth,
    #Inserire la nazione di nascita
    COUNTRY=country_of_birth, 
    #Inserire il giorno di nascita
    DD=day_of_birth, 
    #Inserire l'anno di nascita
    YYYY=year_of_birth,
    #Inserire il nome del corso 
    COURSENAME='Business Strategy & Project Management', 
    #Inserire il mese del corso
    ML='April', 
    #Inserire il mese di rilascio
    MFINAL='April', 
    #Inserire il secondo giorno di lezioni
    D2='15th', 
    #Inserire il giorno di rilascio dell'attestato
    DFINAL='17th', 
    #Inserire il primo giorno di lezioni 
    D1='14th')

document.write('bp_pm.docx')


#Generating Omni-Channel Marketing Certificate
template = "POLIMI.docx"
document = MailMerge(template)
print(document.get_merge_fields())

document.merge(
    #Inserire la città di nascita
    CITY=city, 
    #Inserire il proprio cognome
    SURNAME=surname, 
    #Inserire il proprio nome
    NAME=name, 
    #Inserire il mese di nascita
    MM=month_of_birth,
    #Inserire la nazione di nascita
    COUNTRY=country_of_birth, 
    #Inserire il giorno di nascita
    DD=day_of_birth, 
    #Inserire l'anno di nascita
    YYYY=year_of_birth,
    #Inserire il nome del corso 
    COURSENAME='Omni-Channel Marketing', 
    #Inserire il mese del corso
    ML='May', 
    #Inserire il mese di rilascio
    MFINAL='May', 
    #Inserire il secondo giorno di lezioni
    D2='6th', 
    #Inserire il giorno di rilascio dell'attestato
    DFINAL='8th', 
    #Inserire il primo giorno di lezioni 
    D1='5th')

document.write('omnichannel_marketing.docx')


#Generating ICT Management Certificate
template = "POLIMI.docx"
document = MailMerge(template)
print(document.get_merge_fields())

document.merge(
    #Inserire la città di nascita
    CITY=city, 
    #Inserire il proprio cognome
    SURNAME=surname, 
    #Inserire il proprio nome
    NAME=name, 
    #Inserire il mese di nascita
    MM=month_of_birth,
    #Inserire la nazione di nascita
    COUNTRY=country_of_birth, 
    #Inserire il giorno di nascita
    DD=day_of_birth, 
    #Inserire l'anno di nascita
    YYYY=year_of_birth,
    #Inserire il nome del corso 
    COURSENAME='ICT Management', 
    #Inserire il mese del corso
    ML='June', 
    #Inserire il mese di rilascio
    MFINAL='June', 
    #Inserire il secondo giorno di lezioni
    D2='10th', 
    #Inserire il giorno di rilascio dell'attestato
    DFINAL='12th', 
    #Inserire il primo giorno di lezioni 
    D1='9th')

document.write('ict_management.docx')

print('\n\nCertificates succesfully generated, thank you and have a great day!!!\n\n')