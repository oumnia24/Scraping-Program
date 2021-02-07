import requests
from bs4 import BeautifulSoup
from csv import writer
import xlsxwriter
import xlwt
from xlwt import Workbook

# Workbook is created
wb = Workbook()

# add_sheet is used to create sheet.
sheet1 = wb.add_sheet('Sheet 1')
response = requests.get('https://revues.imist.ma/index.php/JOSSOM')
soup = BeautifulSoup(response.text, 'html.parser')

#Extracting the html code
original_website = requests.get('https://revues.imist.ma/')
original_soup = BeautifulSoup(original_website.text, 'html.parser')


#Getting all the journal names
journal_names= original_soup.find_all('h3')
for i in range(len(journal_names)):
    journal_names[i] = journal_names[i].get_text()
for i in range(len(journal_names)):
    sheet1.write(i, 0, journal_names[i])

#Getting all the journal links
journal_links = original_soup.find_all('a')
journal_links = journal_links[39:]
for i in range(len(journal_links)):
     journal_links[i] = journal_links[i]['href']
journal_links = list(dict.fromkeys(journal_links))

#Adding the full link name:
for i in range(len(journal_links)):
    if journal_links[i][0] == '/':
        complete_link = 'https://revues.imist.ma'
        complete_link += journal_links[i]
        journal_links[i] = complete_link
    sheet1.write(i, 1, journal_links[i])

print(journal_links)
print(len(journal_links))



wb.save('allnames.xls')