import requests
from pprint import pprint
import json
from tabulate import tabulate
import sys
import time
import pandas as pd
import xlsxwriter
import datetime
from pytz import timezone

scholar0 = "0xScholar1"
scholar1 = "0xScholar2"
scholar2 = "0xScholar3"

url = "https://graphql-gateway.axieinfinity.com/graphql"


def GetAxieLatest(address):
# address = scholar1

    url = "https://graphql-gateway.axieinfinity.com/graphql"
    body = {"operationName": "GetAxieLatest",  # to get axies by owner
        "variables": {"from": 0, "size": 100, "sort": "IdDesc", "auctionType": "All", "owner": str(address.strip()),
                      "criteria": {"region": None, "parts": None, "bodyShapes": None, "classes": None, "stages": None,
                                   "numMystic": None, "pureness": None, "title": None, "breedable": None,
                                   "breedCount": None, "hp": [], "skill": [], "speed": [], "morale": []}},
        "query": "query GetAxieLatest($auctionType: AuctionType, $criteria: AxieSearchCriteria, $from: Int, $sort: SortBy, $size: Int, $owner: String) {\n  axies(auctionType: $auctionType, criteria: $criteria, from: $from, sort: $sort, size: $size, owner: $owner) {\n    total\n    results {\n      ...AxieRowData\n      __typename\n    }\n    __typename\n  }\n}\n\nfragment AxieRowData on Axie {\n  id\n  image\n  class\n  name\n  genes\n  owner\n  class\n  stage\n  title\n  breedCount\n  level\n  parts {\n    ...AxiePart\n    __typename\n  }\n  stats {\n    ...AxieStats\n    __typename\n  }\n  auction {\n    ...AxieAuction\n    __typename\n  }\n  __typename\n}\n\nfragment AxiePart on AxiePart {\n  id\n  name\n  class\n  type\n  specialGenes\n  stage\n  abilities {\n    ...AxieCardAbility\n    __typename\n  }\n  __typename\n}\n\nfragment AxieCardAbility on AxieCardAbility {\n  id\n  name\n  attack\n  defense\n  energy\n  description\n  backgroundUrl\n  effectIconUrl\n  __typename\n}\n\nfragment AxieStats on AxieStats {\n  hp\n  speed\n  skill\n  morale\n  __typename\n}\n\nfragment AxieAuction on Auction {\n  startingPrice\n  endingPrice\n  startingTimestamp\n  endingTimestamp\n  duration\n  timeLeft\n  currentPrice\n  currentPriceUSD\n  suggestedPrice\n  seller\n  listingIndex\n  state\n  __typename\n}\n"}


    cont = True #flag to keep trying
    attempts = 0

    while cont:
        if attempts > 0:
            time.sleep(0.5)

        # if attempts == 10:
        #     cont = False

        try:
            attempts += 1
            response = requests.request("POST", url, json=body)
            json_load = json.loads(response.text)
            # print(json_load)
            if "data" in json_load:
                if "results" in json_load["data"]["axies"]:
                    cont = False

        except:
            continue



    response_axies = json_load["data"]["axies"]["results"]

    axie_id = []
    axie_class = []
    axie_parts = []
    axie_parts_eyes = []
    axie_parts_ears = []
    axie_parts_back = []
    axie_parts_mouth = []
    axie_parts_horn = []
    axie_parts_tail = []

    axie_stats = []
    axie_stats_hp = []
    axie_stats_speed = []
    axie_stats_skill = []
    axie_stats_morale = []

    for i in range(json_load["data"]["axies"]["total"]):
        axie_id.append(response_axies[i]["id"])
        axie_class.append(response_axies[i]["class"])
        axie_stats.append(response_axies[i]["stats"])
        axie_stats_hp.append(response_axies[i]["stats"]["hp"])
        axie_stats_speed.append(response_axies[i]["stats"]["speed"])
        axie_stats_skill.append(response_axies[i]["stats"]["skill"])
        axie_stats_morale.append(response_axies[i]["stats"]["morale"])

        del axie_stats[i]["__typename"]

        temp_part_list = []
        for j in range(2,6):
            temp_part_list.append(response_axies[i]["parts"][j]["id"])
            # print("response part is", response_axies[i]["parts"][j]["id"])
            # print(temp_part_list)
        axie_parts.append(temp_part_list)

        axie_parts_eyes.append(response_axies[i]["parts"][0]["id"])
        axie_parts_ears.append(response_axies[i]["parts"][1]["id"])
        axie_parts_back.append(response_axies[i]["parts"][2]["id"][5:])
        axie_parts_mouth.append(response_axies[i]["parts"][3]["id"][6:])
        axie_parts_horn.append(response_axies[i]["parts"][4]["id"][5:])
        axie_parts_tail.append(response_axies[i]["parts"][5]["id"][5:])

    # print(axie_stats_hp, axie_stats_speed, axie_stats_skill, axie_stats_morale)
    axie_stats_compiled = [axie_stats_hp, axie_stats_speed, axie_stats_skill, axie_stats_morale]

    # print(axie_parts_eyes, axie_parts_ears, axie_parts_back, axie_parts_mouth, axie_parts_horn, axie_parts_tail)
    axie_parts_compiled = [axie_parts_eyes, axie_parts_ears, axie_parts_back, axie_parts_mouth, axie_parts_horn, axie_parts_tail]
    return axie_id, axie_class, axie_parts, axie_stats, axie_stats_compiled, axie_parts_compiled
#
# axie_id, axie_class, axie_parts = (GetAxieBriefList(scholar1))

#### axie price

def GetAxiePrice(axie_class, axie_parts, axie_stats):

    hp = [axie_stats["hp"], 61]
    speed = [axie_stats["speed"], 61]
    skill = [axie_stats["skill"], 61]
    morale = [axie_stats["morale"], 61]

    if axie_class == "Beast":
        skill = []
        morale = []
        hp = []
    elif axie_class == "Aquatic":
        skill = []
        morale = []
        hp = []
    elif axie_class == "Plant":
        skill = []
        morale = []
        speed = []


    url = "https://graphql-gateway.axieinfinity.com/graphql"

    body1 = {
  "operationName": "GetAxieLatest",
  "variables": {
    "from": 0,
    "size": 50,
    "sort": "PriceAsc",
    "auctionType": "Sale",
    "criteria": {"classes": axie_class, "parts":axie_parts, "hp": hp, "speed": speed, "skill":skill, "morale":morale}
  },
  "query": "query GetAxieLatest($auctionType: AuctionType, $criteria: AxieSearchCriteria, $from: Int, $sort: SortBy, $size: Int, $owner: String) {\n  axies(auctionType: $auctionType, criteria: $criteria, from: $from, sort: $sort, size: $size, owner: $owner) {\n    total\n    results {\n      ...AxieRowData\n      __typename\n    }\n    __typename\n  }\n}\n\nfragment AxieRowData on Axie {\n  id\n  image\n  class\n  name\n  genes\n  owner\n  class\n  stage\n  title\n  breedCount\n  level\n  parts {\n    ...AxiePart\n    __typename\n  }\n  stats {\n    ...AxieStats\n    __typename\n  }\n  auction {\n    ...AxieAuction\n    __typename\n  }\n  __typename\n}\n\nfragment AxiePart on AxiePart {\n  id\n  name\n  class\n  type\n  specialGenes\n  stage\n  abilities {\n    ...AxieCardAbility\n    __typename\n  }\n  __typename\n}\n\nfragment AxieCardAbility on AxieCardAbility {\n  id\n  name\n  attack\n  defense\n  energy\n  description\n  backgroundUrl\n  effectIconUrl\n  __typename\n}\n\nfragment AxieStats on AxieStats {\n  hp\n  speed\n  skill\n  morale\n  __typename\n}\n\nfragment AxieAuction on Auction {\n  startingPrice\n  endingPrice\n  startingTimestamp\n  endingTimestamp\n  duration\n  timeLeft\n  currentPrice\n  currentPriceUSD\n  suggestedPrice\n  seller\n  listingIndex\n  state\n  __typename\n}\n"

}

    body2 = {
        "operationName": "GetAxieLatest",
        "variables": {
            "from": 0,
            "size": 50,
            "sort": "PriceAsc",
            "auctionType": "Sale",
            "criteria": {"classes": axie_class, "hp": hp, "speed": speed, "skill":skill, "morale":morale}
        },
        "query": "query GetAxieLatest($auctionType: AuctionType, $criteria: AxieSearchCriteria, $from: Int, $sort: SortBy, $size: Int, $owner: String) {\n  axies(auctionType: $auctionType, criteria: $criteria, from: $from, sort: $sort, size: $size, owner: $owner) {\n    total\n    results {\n      ...AxieRowData\n      __typename\n    }\n    __typename\n  }\n}\n\nfragment AxieRowData on Axie {\n  id\n  image\n  class\n  name\n  genes\n  owner\n  class\n  stage\n  title\n  breedCount\n  level\n  parts {\n    ...AxiePart\n    __typename\n  }\n  stats {\n    ...AxieStats\n    __typename\n  }\n  auction {\n    ...AxieAuction\n    __typename\n  }\n  __typename\n}\n\nfragment AxiePart on AxiePart {\n  id\n  name\n  class\n  type\n  specialGenes\n  stage\n  abilities {\n    ...AxieCardAbility\n    __typename\n  }\n  __typename\n}\n\nfragment AxieCardAbility on AxieCardAbility {\n  id\n  name\n  attack\n  defense\n  energy\n  description\n  backgroundUrl\n  effectIconUrl\n  __typename\n}\n\nfragment AxieStats on AxieStats {\n  hp\n  speed\n  skill\n  morale\n  __typename\n}\n\nfragment AxieAuction on Auction {\n  startingPrice\n  endingPrice\n  startingTimestamp\n  endingTimestamp\n  duration\n  timeLeft\n  currentPrice\n  currentPriceUSD\n  suggestedPrice\n  seller\n  listingIndex\n  state\n  __typename\n}\n"

    }

    body3 = {
        "operationName": "GetAxieLatest",
        "variables": {
            "from": 0,
            "size": 20,
            "sort": "PriceAsc",
            "auctionType": "Sale",
            "criteria": {"classes": axie_class}
        },
        "query": "query GetAxieLatest($auctionType: AuctionType, $criteria: AxieSearchCriteria, $from: Int, $sort: SortBy, $size: Int, $owner: String) {\n  axies(auctionType: $auctionType, criteria: $criteria, from: $from, sort: $sort, size: $size, owner: $owner) {\n    total\n    results {\n      ...AxieRowData\n      __typename\n    }\n    __typename\n  }\n}\n\nfragment AxieRowData on Axie {\n  id\n  image\n  class\n  name\n  genes\n  owner\n  class\n  stage\n  title\n  breedCount\n  level\n  parts {\n    ...AxiePart\n    __typename\n  }\n  stats {\n    ...AxieStats\n    __typename\n  }\n  auction {\n    ...AxieAuction\n    __typename\n  }\n  __typename\n}\n\nfragment AxiePart on AxiePart {\n  id\n  name\n  class\n  type\n  specialGenes\n  stage\n  abilities {\n    ...AxieCardAbility\n    __typename\n  }\n  __typename\n}\n\nfragment AxieCardAbility on AxieCardAbility {\n  id\n  name\n  attack\n  defense\n  energy\n  description\n  backgroundUrl\n  effectIconUrl\n  __typename\n}\n\nfragment AxieStats on AxieStats {\n  hp\n  speed\n  skill\n  morale\n  __typename\n}\n\nfragment AxieAuction on Auction {\n  startingPrice\n  endingPrice\n  startingTimestamp\n  endingTimestamp\n  duration\n  timeLeft\n  currentPrice\n  currentPriceUSD\n  suggestedPrice\n  seller\n  listingIndex\n  state\n  __typename\n}\n"

    }

    cont1 = True
    attempts = 0
    while cont1:
        if attempts > 0:
            time.sleep(0.1)

        # if attempts == 10:
        #     cont = False

        try:
            attempts += 1
            response1 = requests.request("POST", url, json=body1)
            json_load1 = json.loads(response1.text)
            if "errors" in json_load1:
                continue
            cont1 = False
            # print("json_load1 is:", json_load1)

            if len(json_load1["data"]["axies"]["results"]) == 0:
                cont2 = True
                attempts = 0
            else:
                cont2 = False
                cont3 = False

        except:
            # attempts += 1
            continue

    json_load2 = {}
    while cont2:
        if attempts > 0:
            time.sleep(0.1)

        # if attempts == 10:
        #     cont = False

        try:
            attempts += 1
            response2 = requests.request("POST", url, json=body2)
            json_load2 = json.loads(response2.text)

            if "errors" in json_load2:
                continue
            cont2 = False
            # print("json_load2 is:", json_load2)

            # print(len(json_load2["data"]["axies"]["results"]))

            if len(json_load2["data"]["axies"]["results"]) == 0:
                cont3 = True
                attempts = 0
            else:
                cont3 = False

        except:
            # attempts += 1
            continue

    json_load3 = {}
    while cont3:
        if attempts > 0:
            time.sleep(0.1)

        # if attempts == 10:
        #     cont = False

        try:
            attempts += 1
            response3 = requests.request("POST", url, json=body3)
            json_load3 = json.loads(response3.text)
            if "errors" in json_load3:
                continue
            cont3 = False
            # print("json_load3 is:", json_load3)
        except:
            # attempts += 1
            continue

    remark = ""

    # print("json_load1", json_load1)     # debug
    # print("json_load2", json_load2)     # debug
    # print("json_load3", json_load3)     # debug

    axieStuck = True
    stuckCounter = 0

    if "data" in json_load1 and len(json_load1["data"]["axies"]["results"]) > 0:     #match class, cards, stats

        while axieStuck:
            if json_load1["data"]["axies"]["results"][stuckCounter]["owner"] != json_load1["data"]["axies"]["results"][stuckCounter]["auction"]["seller"]:
                stuckCounter += 1
            else:
                axieStuck = False

        axie_price_USD = float(json_load1["data"]["axies"]["results"][stuckCounter]["auction"]["currentPriceUSD"])
        axie_price_ETH = float(json_load1["data"]["axies"]["results"][stuckCounter]["auction"]["currentPrice"]) / 1e18

        if axie_class == "Beast":
            remark = "Match cards and speed"
        elif axie_class == "Aquatic":
            remark = "Match cards and speed"
        else:
            remark = "Match cards and all stats"

    elif "data" in json_load2 and len(json_load2["data"]["axies"]["results"]) > 0:       #match class, stats (floor but match stats)
        # print(len(json_load2["data"]["axies"]["results"]))

        while axieStuck:
            if json_load2["data"]["axies"]["results"][stuckCounter]["owner"] != json_load2["data"]["axies"]["results"][stuckCounter]["auction"]["seller"]:
                stuckCounter += 1
            else:
                axieStuck = False

        axie_price_USD = float(json_load2["data"]["axies"]["results"][stuckCounter]["auction"]["currentPriceUSD"])
        axie_price_ETH = float(json_load2["data"]["axies"]["results"][stuckCounter]["auction"]["currentPrice"])/1e18
        remark = "Match stats only"

    elif "data" in json_load3 and len(json_load3["data"]["axies"]["results"]) > 0:                   #match class (zhapalang floor)

        while axieStuck:
            if json_load3["data"]["axies"]["results"][stuckCounter]["owner"] != json_load3["data"]["axies"]["results"][stuckCounter]["auction"]["seller"]:
                stuckCounter += 1
            else:
                axieStuck = False

        axie_price_USD = float(json_load3["data"]["axies"]["results"][stuckCounter]["auction"]["currentPriceUSD"])
        axie_price_ETH = float(json_load3["data"]["axies"]["results"][stuckCounter]["auction"]["currentPrice"]) / 1e18
        remark = "Class floor"

    else:
        axie_price_USD = 0
        axie_price_ETH = 0
        remark = "Price check error"

    return axie_price_ETH, axie_price_USD, remark

def GetScholarName(address):
    url = "https://graphql-gateway.axieinfinity.com/graphql"
    body = {
      "operationName": "GetProfileNameByRoninAddress",
      "variables": {
        "roninAddress": str(address.strip())
      },
      "query": "query GetProfileNameByRoninAddress($roninAddress: String!) {\n  publicProfileWithRoninAddress(roninAddress: $roninAddress) {\n    accountId\n    name\n    __typename\n  }\n}\n"
        }
    # response = requests.request("POST", url, json=body)
    # json_load = json.loads(response.text)
    #
    # # print(json_load) #debug
    # scholarName = json_load["data"]["publicProfileWithRoninAddress"]["name"]
    
    cont = True  # flag to keep trying
    attempts = 0

    while cont:
        if attempts > 0:
            time.sleep(0.5)

        try:
            attempts += 1
            response = requests.request("POST", url, json=body)
            json_load = json.loads(response.text)

            if "errors" not in json_load:
                scholarName = json_load["data"]["publicProfileWithRoninAddress"]["name"]
                cont = False

            # if attempts == 10:
            #     cont = False
        except:
            continue

        time.sleep(0.5)

    return scholarName

def GetTeamTotalValue(address):
    axie_id, axie_class, axie_parts, axie_stats, axie_stats_compiled, axie_parts_compiled = GetAxieLatest(address.strip())

    axie_stats_hp = axie_stats_compiled[0]
    axie_stats_speed = axie_stats_compiled[1]
    axie_stats_skill = axie_stats_compiled[2]
    axie_stats_morale = axie_stats_compiled[3]

    axie_parts_eyes = axie_parts_compiled[0]
    axie_parts_ears = axie_parts_compiled[1]
    axie_parts_back = axie_parts_compiled[2]
    axie_parts_mouth = axie_parts_compiled[3]
    axie_parts_horn = axie_parts_compiled[4]
    axie_parts_tail = axie_parts_compiled[5]

    axie_parts_no_bracket = []
    for i in range(len(axie_parts)):
        axie_parts_no_bracket.append(str(axie_parts[i])[1:-1])


    axie_stats_no_bracket = []
    for i in range(len(axie_stats)):
        axie_stats_no_bracket.append(", ".join(str(key) + ": " + str(value) for key, value in axie_stats[i].items()))


    axie_price_USD_list = []
    axie_price_ETH_list = []
    remarks_list = []
    for i in range(len(axie_id)):
        # print("Now parsing Axie ID:", axie_id[i]) #debug
        axie_price_ETH, axie_price_USD, remark = GetAxiePrice(axie_class[i], axie_parts[i], axie_stats[i])
        axie_price_USD_list.append(round(axie_price_USD,3))
        axie_price_ETH_list.append(round(axie_price_ETH,3))
        remarks_list.append(remark)

    table = {"Axie ID" : axie_id, "Class":axie_class , "Parts" : axie_parts_no_bracket , "Stats": axie_stats_no_bracket,"Price (ETH)" : axie_price_ETH_list, "Price (USD)" : axie_price_USD_list, "Price remark" : remarks_list}
    print(tabulate(table, headers="keys", tablefmt="fancy_grid"))

    scholarName = GetScholarName(address.strip())
    placeholder_name = [scholarName]
    name_list = placeholder_name * len(axie_id)
    if len(axie_id) == 0:
        name_list = [scholarName]
        axie_id = [""]
        axie_class = [""]
        axie_parts_no_bracket = [""]
        axie_stats_no_bracket = [""]
        axie_price_ETH_list = [0]
        axie_price_USD_list = [0]
        remarks_list = ["Empty account"]

        axie_stats_hp = [""]
        axie_stats_speed = [""]
        axie_stats_skill = [""]
        axie_stats_morale = [""]

        axie_parts_eyes = [""]
        axie_parts_ears = [""]
        axie_parts_back = [""]
        axie_parts_mouth = [""]
        axie_parts_horn = [""]
        axie_parts_tail = [""]

    else:
        axie_stats_hp = axie_stats_compiled[0]
        axie_stats_speed = axie_stats_compiled[1]
        axie_stats_skill = axie_stats_compiled[2]
        axie_stats_morale = axie_stats_compiled[3]

        axie_parts_eyes = axie_parts_compiled[0]
        axie_parts_ears = axie_parts_compiled[1]
        axie_parts_back = axie_parts_compiled[2]
        axie_parts_mouth = axie_parts_compiled[3]
        axie_parts_horn = axie_parts_compiled[4]
        axie_parts_tail = axie_parts_compiled[5]

    table2 = {"Scholar": name_list, "Axie ID": axie_id, "Class": axie_class, "Parts": axie_parts_no_bracket, "Stats": axie_stats_no_bracket, "Price (ETH)": axie_price_ETH_list, "Price (USD)": axie_price_USD_list, "Price remark": remarks_list, "Back": axie_parts_back, "Mouth": axie_parts_mouth, "Horn": axie_parts_horn, "Tail": axie_parts_tail, "HP": axie_stats_hp, "Speed": axie_stats_speed, "Skill": axie_stats_skill, "Morale": axie_stats_morale}

    print("Total current selling price of", scholarName + " "+address.strip(), "team is USD", round(sum(axie_price_USD_list),3))
    print("\n")
    return sum(axie_price_USD_list), sum(axie_price_ETH_list), table2

def main():
    tick = time.time()
    timenow = datetime.datetime.now(timezone("Asia/Kuala_Lumpur"))

    filename_date_timestamp = f" {timenow.year}{timenow.month}{timenow.day:02} {timenow.hour:02}{timenow.minute:02}"
    filenameTXT = fileOwner + filename_date_timestamp + ".txt"
    filenameXLSX = fileOwner + filename_date_timestamp + ".xlsx"

    if printfile:
    # with open(filenameTXT,'wt', encoding="utf-8") as file:
        file = open(filenameTXT,'wt', encoding="utf-8")
        sys.stdout = file

    running_total = 0
    running_total_ETH = 0
    overall_list = {"Scholar": [], "Axie ID": [], "Class": [], "Parts": [], "Stats": [], "Price (ETH)": [], "Price (USD)": [], "Price remark": [], "Back": [], "Mouth": [], "Horn": [], "Tail": [], "HP": [], "Speed": [], "Skill": [], "Morale": []}
    for i in list_of_scholar:

        current_team_value, current_team_value_ETH, current_team = GetTeamTotalValue(i)

        running_total = round(running_total + current_team_value, 3)
        running_total_ETH = round(running_total_ETH + current_team_value_ETH, 3)
        for key, value in overall_list.items():
            value.extend(current_team[key])

        # pprint(overall_list)


        # try:
        #     running_total = round(running_total + GetTeamTotalValue(i), 3)
        #
        # except requests.exceptions.ConnectionError:
        #     time.sleep(0.5)
        #     running_total = round(running_total + GetTeamTotalValue(i), 3)
        #
        # except:
        #     time.sleep(0.5)
        #     running_total = round(running_total + GetTeamTotalValue(i), 3)
        # time.sleep(0.5)

    print("Total portfolio selling price is " + str(running_total_ETH) + " ETH, or USD " + str(running_total) + ". Total recoverable cost (after 4.25% marketplace tax) is USD " + str(round(running_total*0.9575, 2)))

    tock = time.time()
    ticktock = tock-tick

    # timenow = datetime.datetime.now(timezone("Asia/Kuala_Lumpur"))
    month = timenow.strftime("%b")
    print(f"Total time taken for this query of {len(list_of_scholar):.0f} scholars: {ticktock:.2f} seconds")
    print(f"Query requested by {fileOwner}, which contains {len(list_of_scholar)} scholars/addresses.")
    print(f"Query completed on {timenow.day} {month} {timenow.year} at {timenow.hour:02}{timenow.minute:02}hrs UTC+8 (Malaysia Time)")
    if printfile:
        file.close()

    sys.stdout = sys.__stdout__

    if printfile:

        with open(filenameTXT, "r", encoding="utf-8") as file2:
                print(file2.read())

    # pprint(overall_list)

        workbook = xlsxwriter.Workbook(filenameXLSX)
        worksheet1 = workbook.add_worksheet()
        cell_format = workbook.add_format()
        workbook.formats[0].set_font_name('Arial')
        workbook.formats[0].set_font_size(10)
        worksheet1.write_string('A1', "Scholar")
        worksheet1.write_column('A2', overall_list["Scholar"])
        worksheet1.write_string('B1', "Axie ID")
        worksheet1.write_column('B2', overall_list["Axie ID"])
        worksheet1.write_string('C1', "Class")
        worksheet1.write_column('C2', overall_list["Class"])
        # worksheet1.write_string('D1', "Parts")
        # worksheet1.write_column('D2', overall_list["Parts"])
        worksheet1.write_string('D1', "Back")
        worksheet1.write_column('D2', overall_list["Back"])
        worksheet1.write_string('E1', "Mouth")
        worksheet1.write_column('E2', overall_list["Mouth"])
        worksheet1.write_string('F1', "Horn")
        worksheet1.write_column('F2', overall_list["Horn"])
        worksheet1.write_string('G1', "Tail")
        worksheet1.write_column('G2', overall_list["Tail"])
        # worksheet1.write_string('H1', "Stats")
        # worksheet1.write_column('H2', overall_list["Stats"])
        worksheet1.write_string('H1', "HP")
        worksheet1.write_column('H2', overall_list["HP"])
        worksheet1.write_string('I1', "Speed")
        worksheet1.write_column('I2', overall_list["Speed"])
        worksheet1.write_string('J1', "Skill")
        worksheet1.write_column('J2', overall_list["Skill"])
        worksheet1.write_string('K1', "Morale")
        worksheet1.write_column('K2', overall_list["Morale"])
        worksheet1.write_string('L1', "Price (ETH)")
        worksheet1.write_column('L2', overall_list["Price (ETH)"])
        worksheet1.write_string('M1', "Price (USD)")
        worksheet1.write_column('M2', overall_list["Price (USD)"])
        worksheet1.write_string('N1', "Remark")
        worksheet1.write_column('N2', overall_list["Price remark"])
        worksheet1.set_column(0, 0, 20)
        worksheet1.set_column(3, 6, 15)
        worksheet1.set_column(7, 10, 6)
        worksheet1.set_column(11, 12, 11)
        worksheet1.set_column(13, 13, 30)

        workbook.close()

if __name__ == "__main__":

    fileOwner = "OWNERNAME" #free text
    list_of_scholar = [scholar1, scholar2, scholar0]  # own

    # printfile = False
    printfile = True



    main()
