import requests
from bs4 import BeautifulSoup
from csv import writer
import xlwt
from xlwt import Workbook
import xlrd


#Creating a new sheet
wb = Workbook()
sheet1 = wb.add_sheet('ISSNs')

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

#scraping the comite editorial for directeur de publication
for i in range(len(list_of_journal_links_clean)):
    try:
        sheet1.write(i, 0, list_of_journal_links_clean[i])
        response = requests.get(list_of_journal_links_clean[i])
        soup = BeautifulSoup(response.text, 'html.parser')
        element = soup.find(id='navItem-0').find('a')
        element = element['href']
        access_element = requests.get(element)
        opened_element = BeautifulSoup(access_element.text, 'html.parser')
        directeur = opened_element.p
        print(directeur)
        sheet1.write(i , 1, directeur.get_text())
    except AttributeError:
        print("skipped")


# i= 1
# for link in list_of_journal_links:
#      try:
#         response = requests.get(link)
#         soup = BeautifulSoup(response.text, 'html.parser')
#         element = soup.find(id='archives').find('a')
#         element = element['href']
#         access_element = requests.get(element)
#         opened_element = BeautifulSoup(access_element.text, 'html.parser')
#         issues_info = opened_element.find(id ='issues').find_all('h3')
#         for j in range(len(issues_info)):
#             issues_info[j] = issues_info[j].get_text()
#             issues_info[j] = int(issues_info[j])
#         issues_info.sort()
#         if 2019 in issues_info:
#             archive_2019 = opened_element.find(id='issues').find_all('h4')
#             links_2019 = []
#             for header in archive_2019:
#                 if '2019' in header.get_text():
#                     links_2019.append(header.find('a')['href'])
#             number_of_articles_2019 = 0
#             for link in links_2019:
#                 response = requests.get(link)
#                 soup = BeautifulSoup(response.text, 'html.parser')
#                 annoying_cover_image = soup.find(id='issueCoverImage')
#                 if annoying_cover_image:
#                     link = annoying_cover_image.find('a')['href']
#                     response = requests.get(link)
#                     soup = BeautifulSoup(response.text, 'html.parser')
#                 article_links = soup.find_all(class_='tocArticle')
#                 number_of_articles_2019 += len(article_links)
#             sheet1.write(i - 1, 0, number_of_articles_2019)
#             print(number_of_articles_2019)
#      except AttributeError:
#          print("nothing in 2019")
#          sheet1.write(i-1, 0, "nothing in 2019")
#      i +=1



    # opened_2019_archive = requests.get(archive_2019_link)
    # soup = BeautifulSoup(opened_2019_archive.text, 'html.parser')
    # print("Number of articles in 2019:", len(soup.find_all(class_='tocArticle')) -1)
    # sheet1.write(i,0, len(soup.find_all(class_='tocArticle')) -1)
    # except AttributeError:
    #     print("this gave an error. try again")
    # i +=1

# sheet1.write(0,3, contact_info_email)
#
# #Accessing old and new volumes
# issues_info = opened_archives.find(id ='issues').find_all('h3')
# for i in range(len(issues_info)):
#     issues_info[i] = issues_info[i].get_text()
#     issues_info[i] = int(issues_info[i])
# issues_info.sort()
# oldest_volume = issues_info[0]
# newest_volume = issues_info[len(issues_info) - 1]
# print('Oldest volume:', oldest_volume)
# # sheet1.write(0,4, oldest_volume)
# print('Current volume:', newest_volume)
# # sheet1.write(0,5, newest_volume)
#
# #Getting the number of articles in 2019
# if 2019 in issues_info:
#     archive_2019 = opened_archives.find(id='issues').find('h3', string = '2019')
#     # print(archive_2019.next_element.next_element.next_element.find('a')['href'])
#     archive_2019_link = archive_2019.next_element.next_element.next_element.find('a')['href']
#     opened_2019_archive = requests.get(archive_2019_link)
#     soup = BeautifulSoup(opened_2019_archive.text, 'html.parser')
#     print("Number of articles in 2019:", len(soup.find_all(class_='tocArticle')) -1)
#     # sheet1.write(0,6, len(soup.find_all(class_='tocArticle')) -1)
#
# #Getting the review info
# policies_links = opened_a_propos.find(id= 'aboutPolicies')
# politique_de_rubrique = policies_links.find('a', string = 'Politiques de rubriques')['href']
# politique_de_rubrique = requests.get(politique_de_rubrique)
#
# politique_de_rubrique = BeautifulSoup(politique_de_rubrique.text, 'html.parser')
# list_of_policies = politique_de_rubrique.find(id='sectionPolicies').find_all('td')
# if list_of_policies[5].text == " Évalué par les pairs":
#     print("Review info: Yes")
#
# else:
#     print("Review info: No")
#
#Accessing contact information
        # contact_section= opened_a_propos.find(id = 'aboutPeople').find('ul').find('li').find('a')
        # contact_section = contact_section['href']
        # contact_section = requests.get(contact_section)
        # contact_section = BeautifulSoup(contact_section.text, 'html.parser')
        # contact_info = contact_section.find(id = 'principalContact')
        # contact_info_name = contact_info.find('strong')
        # contact_info_name = contact_info_name.get_text()
        # contact_info_email = contact_info.find('a')
        # contact_info_email = contact_info_email.get_text()
        # print('Contact person:', contact_info_name)
        # # sheet1.write(0, 2, contact_info_name)
        # print('Contact email:', contact_info_email)

     #     #Finding ISSN:
     #     ISSN_info = soup.find(id = 'pageFooter').get_text()
     #     print(i, 'ISSN:', ISSN_info[6:])
     #     sheet1.write(i -1, 0, ISSN_info[6:])
     # except AttributeError:
     #     print(i, "Oops!  That was no valid number.  Try again...")
     #     sheet1.write(i - 1, 0, 'oops: no valid number')
     # i+=1

     # Accessing elements of the menu:
     # def access_element(id_info):
     #     element = soup.find(id=id_info).find('a')
     #     element = element['href']
     #     access_element = requests.get(element)
     #     opened_element = BeautifulSoup(access_element.text, 'html.parser')
     #     return opened_element


     # Accessing A propos
     # opened_a_propos = access_element('about')
#
# Accessing Equipe Editoriale
# opened_equipe_editoriale = access_element('navItem-0')

# Accessing Archives
# opened_archives = access_element('archives')

wb.save('number of articles.xls')