"""
Module for verifying landing page URLs in advertisements on the TSmedia (Graphite Adserver) platform.
Alerts whenever response != 200 for a given URL.
"""
import requests
from bs4 import BeautifulSoup as bs
from datetime import datetime
from collections import OrderedDict

url_main  = "https://graphite.tsmedia.si"
url_login = "http://graphite.tsmedia.si/dashboard/login"
url_data  = "https://graphite.tsmedia.si/dashboard/campaigns-all"
payload = {"email": "damjan.mihelic@tsmedia.si", "password": "L!4)ggdSkRH-N/Sqe9Qq-/379YbO4X"}
response_errors = OrderedDict()
with requests.Session() as s:
    login = s.post(url_login, data=payload)
    request_main = s.get(url_data)
    soup_main = bs(request_main.content, "html5lib")
    campaigns = soup_main.find_all("tr")
    for campaign in campaigns:
        try: # IndexError skips first element in list, which is an empty list
            if campaign.find_all("td", class_="right")[0].getText() == "✔": # gets only active campaigns (not ✘)
                url_campaign = url_main + campaign.find_all("a", class_="button small outline")[0].get("href")
                request_campaign = s.get(url_campaign)
                soup_campaign = bs(request_campaign.content, "html5lib")
                ads_outline = soup_campaign.find_all("tr")[1:]
                ads = list()
                for ad_outline in ads_outline:
                    if ad_outline.find_all("td", class_="center")[0].getText() == "✔": # gets only active ads (not ✘)
                        ads.append(ad_outline.find_all("td", class_="right")[0])
                for ad in ads:
                    url_ad = url_main + ad.find_all("a", class_="button small outline")[0].get("href")
                    request_ad = s.get(url_ad)
                    soup_ad = bs(request_ad.content, "html5lib")
                    url_landing = soup_ad.find_all("input", id="final_url")[0].get("value")
                    print(url_landing)
                    try:
                        counter = 0
                        while counter < 5: # trying
                            status = requests.get(url_landing).status_code
                            if status != 200: counter += 1
                            else: break
                        if status != 200:
                            response_errors[url_landing] = status
                    except TimeoutError: response_errors[url_landing] = "TimeoutErr"
                    except ConnectionError: response_errors[url_landing] = "ConnectionErr"
                    except requests.exceptions.MissingSchema: response_errors[url_landing] = "InvalidUrlErr"
                    except requests.exceptions.SSLError: response_errors[url_landing] = "SSLCertErr"
        except IndexError: continue
if response_errors:
    fh = open("URL_errors.txt", "a")
    fh.write("URL errors fetched on {}\n".format(str(datetime.now())[:19]))
    for url, error in response_errors.items(): fh.write(url + ": " + str(error) + "\n")
    fh.write("\n")
    fh.close()
    print("URL errors saved in file URL_errors.txt.")