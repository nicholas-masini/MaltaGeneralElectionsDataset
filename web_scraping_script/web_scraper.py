from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from openpyxl import load_workbook

# absolute path
s = Service("C:/Users/user/Desktop/testing/chromedriver")
op = webdriver.ChromeOptions()

# List of canditates
canditates = []

webdriver = webdriver.Chrome(
    service = s,
    options = op
)

with webdriver as driver:
    print("\nStarting web scraping.\n")

    election_token_17 = "8NDVTyBTAoGgFAux3fFn-Ed8AHsvJBT5wC18Yo8qS4ZBMqJVIvKu6M_cYEQ1ldECVgvYXsX6EWCXa_N-dIXjZagG8o--uwINgzYRVVBu3yA1"
    election_token_22 = "mhmg8-e5kHQPlSASCxGVJFqLulzJIAkxcQThtlteXTpVlBc5MQ84suA-Pcem_cmKyk-TOqRX4J7ZfuA8uE4nVW71WxAvxIszdjZfn0yiP-A1"

    for election in range(244, 249, 4):

        token = ""
        upper, lower, year = 0, 0, 0
        if election == 244: 
            token = election_token_17
            upper = 777
            lower = 790
            year = 2017
        elif election == 248:
            token = election_token_22
            upper = 877
            lower = 890
            year = 2022

        # Iterating through every district
        for district_index in range(upper, lower):

            # URL to scrape data from
            url = "https://electoral.gov.mt/ElectionResults/ElecResults/ZoneCriteriaChanged?SelectedZone="+str(district_index)+"&SelectedElection="+str(election)+"&ElectionType=1&__RequestVerificationToken="+token

            # timeout
            wait = WebDriverWait(driver, 10)

            print("District: "+str(district_index-776))

            # retrieve data
            driver.get(url)

            if(district_index == 777 or district_index == 877 or district_index == 884 or district_index == 888 or district_index == 889): div_id = 3
            else: div_id = 4
            
            # Obtain the number of rows in body
            rows = 1+len(driver.find_elements(By.XPATH, "/html/body/div[3]/div["+str(div_id)+"]/table/tbody/tr"))

            # Obtain the number of coloumns in table
            cols = len(driver.find_elements(By.XPATH, "/html/body/div[3]/div["+str(div_id)+"]/table/tbody/tr[1]/td"))

            # Obtain some district results information
            registered_voters = int((driver.find_element(By.XPATH, "/html/body/div[2]/div[2]/div[1]/div[2]/div[3]/div").text).replace(',', ''))
            valid_votes = int((driver.find_element(By.XPATH, "/html/body/div[2]/div[2]/div[1]/div[2]/div[5]/div[2]").text).replace(',', ''))
            quota = int((driver.find_element(By.XPATH, "/html/body/div[2]/div[2]/div[1]/div[2]/div[2]/div").text).replace(',', ''))


            ct1_sum_PL = 0
            ct1_sum_PN = 0
            ct1_sum_AD = 0
            ct1_sum_MPM = 0
            ct1_sum_AB = 0
            ct1_sum_PP = 0
            ct1_sum_V = 0

            # Getting "PTotal" attribute 
            for r in range(1, rows):
                # Getting party
                party = driver.find_element(By.XPATH, "/html/body/div[3]/div["+str(div_id)+"]/table/tbody/tr["+str(r)+"]/td["+str(2)+"]").text

                if(party == "Partit Laburista"): ct1_sum_PL += int((driver.find_element(By.XPATH, "/html/body/div[3]/div["+str(div_id)+"]/table/tbody/tr["+str(r)+"]/td["+str(4)+"]").text).replace(',', ''))
                elif(party == "Partit Nazzjonalista"): ct1_sum_PN += int((driver.find_element(By.XPATH, "/html/body/div[3]/div["+str(div_id)+"]/table/tbody/tr["+str(r)+"]/td["+str(4)+"]").text).replace(',', ''))
                elif(party == "Alternattiva Demokratika" or party == "AD + PD"): ct1_sum_AD += int((driver.find_element(By.XPATH, "/html/body/div[3]/div["+str(div_id)+"]/table/tbody/tr["+str(r)+"]/td["+str(4)+"]").text).replace(',', ''))
                elif(party == "Moviment Patrijotti Maltin"): ct1_sum_MPM += int((driver.find_element(By.XPATH, "/html/body/div[3]/div["+str(div_id)+"]/table/tbody/tr["+str(r)+"]/td["+str(4)+"]").text).replace(',', ''))
                elif(party == "Alleanza Bidla" or party == "ABBA"): ct1_sum_AB += int((driver.find_element(By.XPATH, "/html/body/div[3]/div["+str(div_id)+"]/table/tbody/tr["+str(r)+"]/td["+str(4)+"]").text).replace(',', ''))
                elif(party == "Partit Popolari"): ct1_sum_PP += int((driver.find_element(By.XPATH, "/html/body/div[3]/div["+str(div_id)+"]/table/tbody/tr["+str(r)+"]/td["+str(4)+"]").text).replace(',', ''))
                elif(party == "Volt Malta"): ct1_sum_V += int((driver.find_element(By.XPATH, "/html/body/div[3]/div["+str(div_id)+"]/table/tbody/tr["+str(r)+"]/td["+str(4)+"]").text).replace(',', ''))

            # Iterating through every canditate
            for r in range(1, rows):

                canditate = {}

                # Getting name 
                name = driver.find_element(By.XPATH, "/html/body/div[3]/div["+str(div_id)+"]/table/tbody/tr["+str(r)+"]/td["+str(1)+"]").text
                canditate["NAME"] = name
                print("Canditate: "+name)

                # Getting party
                party = driver.find_element(By.XPATH, "/html/body/div[3]/div["+str(div_id)+"]/table/tbody/tr["+str(r)+"]/td["+str(2)+"]").text
                # Assinging party value according to general elections data set codebook
                if(party == "Partit Laburista"): 
                    canditate["PARTY"] = 13
                    canditate["PTOTAL"] = ct1_sum_PL
                elif(party == "Partit Nazzjonalista"): 
                    canditate["PARTY"] = 12
                    canditate["PTOTAL"] = ct1_sum_PN
                elif(party == "Alternattiva Demokratika" or party == "AD + PD"): 
                    canditate["PARTY"] = 28
                    canditate["PTOTAL"] = ct1_sum_AD
                elif(party == "Moviment Patrijotti Maltin"): 
                    canditate["PARTY"] = 36
                    canditate["PTOTAL"] = ct1_sum_MPM
                elif(party == "Alleanza Bidla" or party == "ABBA"): 
                    canditate["PARTY"] = 37
                    canditate["PTOTAL"] = ct1_sum_AB
                elif(party == "Partit Popolari"):
                    canditate["PARTY"] = 38
                    canditate["PTOTAL"] = ct1_sum_PP
                elif(party == "Volt Malta"):
                    canditate["PARTY"] = 39
                    canditate["PTOTAL"] = ct1_sum_V
                elif(party == "Independent Candidate" or party == "Kandidat Indipendenti"): canditate["PARTY"] = 50

                canditate["counts"] = []
                # Getting counts
                for c in range(4, cols+1, 2):
                    count = driver.find_element(By.XPATH, "/html/body/div[3]/div["+str(div_id)+"]/table/tbody/tr["+str(r)+"]/td["+str(c)+"]").text
                    if(count == "..."): canditate["counts"].append(0)
                    elif count == '': pass
                    else: canditate["counts"].append(int(count.replace(',', '')))

                if(party == "Independent Candidate" or party == "Kandidat Indipendenti"): canditate["PTOTAL"] = max(canditate["counts"])

                # Setting other attributes
                canditate["YEAR"] = year
                if year == 2017: canditate["Dist"] = district_index-776
                else: canditate["Dist"] = district_index-876
                canditate["REG"] = registered_voters
                canditate["VOTE"] = valid_votes
                canditate["QUOTA"] = quota
                canditate["COUNTS"] = len(canditate["counts"])
                canditate["PSIZE"] = (canditate["PTOTAL"] / valid_votes) * 100
                canditate["PSHARE"] = (canditate["counts"][0] / canditate["PTOTAL"]) * 100
                canditate["QSHARE"] = (canditate["counts"][0] / quota) * 100
                canditate["TOPS"] = max(canditate["counts"])
                canditate["LAST"] = canditate["counts"][len(canditate["counts"])-1]

                canditates.append(canditate)

            # Getting non-transferrable votes
            non_trans = {}

            non_trans["NAME"] = "* Non-Trans. *"
            non_trans["PARTY"] = 99
            non_trans["counts"] = []
            for c in range(4, cols+1, 2):
                count = driver.find_element(By.XPATH, "/html/body/div[3]/div["+str(div_id)+"]/table/tfoot/tr[1]/th["+str(c)+"]").text
                if(count != "..." and count != ''): non_trans["counts"].append(int(count.replace(',', '')))
            non_trans["PTOTAL"] = 0
            non_trans["YEAR"] = year
            if year == 2017: non_trans["Dist"] = district_index-776
            else: non_trans["Dist"] = district_index-876
            non_trans["REG"] = registered_voters
            non_trans["VOTE"] = valid_votes
            non_trans["QUOTA"] = quota
            non_trans["COUNTS"] = len(non_trans["counts"])
            non_trans["PSIZE"] = 0
            non_trans["PSHARE"] = 0
            non_trans["QSHARE"] = 0
            non_trans["TOPS"] = max(non_trans["counts"])
            non_trans["LAST"] = non_trans["counts"][len(non_trans["counts"])-1]

            canditates.append(non_trans)
            print()
        
    driver.close()

print()
print("Appending data to existing dataset.")

# Appending new data to existing dataset
book = load_workbook('filtered_dataset.xlsx')

worksheet = book.active

for c in canditates:
    # Adding 'n/a' counts
    row = [
            c["NAME"],
            c["YEAR"],
            c["Dist"],
            c["REG"],
            c["VOTE"],
            c["QUOTA"],
            c["COUNTS"],
            c["PTOTAL"],
            c["PSIZE"],
            c["PARTY"],
            c["QSHARE"],
            c["PSHARE"],
            c["TOPS"],
            c["LAST"],
          ]
    for i in range(len(c["counts"]), 37): c["counts"].append("n/a")
    for count in c["counts"]: row.append(count)
    worksheet.append(row)

book.save('new_filtered_dataset.xlsx')

print("Finished.")