import requests
from bs4 import BeautifulSoup
from csv import writer
import xlwt
from xlwt import Workbook
import xlrd


#Creating a new sheet
wb = Workbook()
sheet1 = wb.add_sheet('Journal Info')

#Extracting the links from previous worksheet
workbook = xlrd.open_workbook('/Users/oumniachellah/Downloads/Relevantlinks.xlsx')
worksheet = workbook.sheet_by_index(0)
list_of_journal_links = []
for i in range(136):
    list_of_journal_links.append(worksheet.cell(i, 0).value)

#Removing the non imist journals
list_of_journal_links_clean = []
for i in range(len(list_of_journal_links)):
    if "imist" in list_of_journal_links[i]:
        list_of_journal_links_clean.append(list_of_journal_links[i])
print(list_of_journal_links_clean)
list_of_journal_links_clean.remove('https://revues.imist.ma/?journal=Organisation-Territoires')


for i in range(len(list_of_journal_links_clean)):
    try:
        sheet1.write(i, 0, list_of_journal_links_clean[i])
        response = requests.get(list_of_journal_links_clean[i])
        soup = BeautifulSoup(response.text, 'html.parser')

        # Opening the comite editorial section, and scraping it for directeur de publication
        comite = soup.find(id='navItem-0').find('a')
        comite = comite['href']
        comite_link = requests.get(comite)
        opened_comite = BeautifulSoup(comite_link.text, 'html.parser')
        directeur = opened_comite.p
        sheet1.write(i , 1, directeur.get_text())
        #Finding ISSN:
        ISSN_info = soup.find(id = 'pageFooter').get_text()
        sheet1.write(i, 2, ISSN_info[6:])
        # Accessing A propos
        a_propos = soup.find(id='navItem-0').find('a')
        a_propos = a_propos['href']
        a_propos_link = requests.get(a_propos)
        opened_a_propos = BeautifulSoup(a_propos_link.text, 'html.parser')
        # Accessing contact information within a_propos: Opening contact section, getting main contact's name and email
        contact_section= opened_a_propos.find(id = 'aboutPeople').find('ul').find('li').find('a')
        contact_section = contact_section['href']
        contact_section = requests.get(contact_section)
        contact_section = BeautifulSoup(contact_section.text, 'html.parser')
        contact_info = contact_section.find(id = 'principalContact')
        contact_info_name = contact_info.find('strong')
        contact_info_name = contact_info_name.get_text()
        contact_info_email = contact_info.find('a')
        contact_info_email = contact_info_email.get_text()
        sheet1.write(i, 3, contact_info_name)
        sheet1.write(i, 4, contact_info_email)
        #Getting the review info
        policies_links = opened_a_propos.find(id= 'aboutPolicies')
        politique_de_rubrique = policies_links.find('a', string = 'Politiques de rubriques')['href']
        politique_de_rubrique = requests.get(politique_de_rubrique)
        politique_de_rubrique = BeautifulSoup(politique_de_rubrique.text, 'html.parser')
        list_of_policies = politique_de_rubrique.find(id='sectionPolicies').find_all('td')
        if list_of_policies[5].text == " Évalué par les pairs":
            sheet1.write(i, 5, "Yes")
        else:
            sheet1.write(i, 5, "No")
    except AttributeError:
        sheet1.write(i, 1, "No Directeur Info")
        sheet1.write(i, 2, "No ISSN Info")

#Opening the archives section, opening the 2019 edition, counting the number of articles.
for i in range(len(list_of_journal_links)):
     try:
        link = list_of_journal_links[i]
        response = requests.get(link)
        soup = BeautifulSoup(response.text, 'html.parser')
        archives = soup.find(id='archives').find('a')
        archives = archives['href']
        archives_link = requests.get(archives)
        opened_archives = BeautifulSoup(archives_link.text, 'html.parser')
        issues_info = opened_archives.find(id ='issues').find_all('h3')
        for j in range(len(issues_info)):
            issues_info[j] = issues_info[j].get_text()
            issues_info[j] = int(issues_info[j])
        issues_info.sort()
        #Getting oldest and newest issue dates:
        oldest_volume = issues_info[0]
        newest_volume = issues_info[len(issues_info) - 1]
        sheet1.write(0, 3, oldest_volume)
        sheet1.write(0, 4, newest_volume)
        #Number of articles in 2019
        if 2019 in issues_info:
            archive_2019 = opened_archives.find(id='issues').find_all('h4')
            links_2019 = []
            for header in archive_2019:
                if '2019' in header.get_text():
                    links_2019.append(header.find('a')['href'])
            number_of_articles_2019 = 0
            for link in links_2019:
                response = requests.get(link)
                soup = BeautifulSoup(response.text, 'html.parser')
                cover_image = soup.find(id='issueCoverImage')
                if cover_image:
                    link = cover_image.find('a')['href']
                    response = requests.get(link)
                    soup = BeautifulSoup(response.text, 'html.parser')
                article_links = soup.find_all(class_='tocArticle')
                number_of_articles_2019 += len(article_links)
            sheet1.write(i, 2, number_of_articles_2019)
     except AttributeError:
         print("nothing in 2019")
         sheet1.write(i, 2, "nothing in 2019")



wb.save('Journal Data.xls')