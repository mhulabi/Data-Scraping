import xlsxwriter
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.support import expected_conditions as EC
from time import sleep
cities = []
spaces = []

name = []
location = []
description = []
address = []
facebook = []
website = []
rates = []
pOffice = []
dDesk = []
mRoom = []
hDesk = []
hour = []
amenity = []

kingUrl = f"https://www.coworkbooking.com/europe/italy"
driver = webdriver.Chrome()
driver.maximize_window()
driver.get(kingUrl)
sleep(15)
stop = 0
sleep(2)
html = driver.page_source
soup = BeautifulSoup(html, "html.parser")
parent = soup.find_all("li", class_="city")
for item in parent:
    cityh = item.find("a", class_="show_all_coworks")
    cities.append("https://www.coworkbooking.com" + cityh['href'])
print(cities)

for city in cities:
    driver.get(city)
    sleep(2)
    sleep(2)
    html = driver.page_source
    soup = BeautifulSoup(html, "html.parser")
    work = soup.find_all("li", class_="space")
    for item2 in work:
        #stop += 1
        # if stop == 5:
        # break
        space = item2.find("a", class_="cta noblink")
        spaces.append("https://www.coworkbooking.com" + space['href'])
    # if stop == 5:
        # break

page = 1
count = 0
index = 1
print(spaces)
for s in spaces:
    print(count)
    count += 1
    driver.get(s)
    sleep(2)
    # sleep(2)
    html = driver.page_source
    soup = BeautifulSoup(html, "html.parser")

    # Name
    try:
        n = soup.find("h1")
        name.append(n.text)
    except AttributeError:
        location.append("")

    # City
    try:
        c = soup.find("div", class_="parent_place")
        location.append(c.text.replace("Europe / Italy / ", ""))
    except AttributeError:
        location.append("")

    # Description
    descs = soup.find("div", id="space_description")
    try:
        description.append(descs.text)  # .strip("\n").strip(" "))
    except AttributeError:
        description.append("")
    # Rates
    rates = soup.find(
        "ul", class_="offers")
    try:
        prices = rates.find_all("li", class_="article")
        confirm = [False, False, False, False]
        pText = ""
        hText = ""
        dText = ""
        mText = ""
        for rate in prices:
            # Private Office
            if "Private Office" in rate.text:
                pDetail = rate.find("span", class_="price")
                sDetail = rate.find("span", class_="article_type")
                confirm[0] = True
                try:
                    pDetail2 = pDetail.find("span", class_="price_original")
                    pText += ("Price: " +
                              pDetail2.text.replace("31 days", "month").replace("29 days", "month").replace("month", "/month").replace("day", "/day").replace("hour", "/hour"))
                except AttributeError:
                    pText += ("Price: " +
                              pDetail.text.replace("31 days", "month").replace("29 days", "month").replace("month", "/month").replace("day", "/day").replace("hour", "/hour"))
                ssplit = sDetail.text.split("people ")
                try:
                    pText += (", Capacity: " +
                              ssplit[0][-3:].strip(" ") + " people | ")
                except IndexError:
                    pText += (", Capacity: " + sDetail.text + " | ")
            # Hot Desk
            elif "Hot Desk" in rate.text:
                pDetail = rate.find("span", class_="price")
                sDetail = rate.find("span", class_="article_type")
                confirm[1] = True
                try:
                    pDetail2 = pDetail.find("span", class_="price_original")
                    hText += ("Price: " +
                              pDetail2.text.replace("31 days", "month").replace("29 days", "month").replace("month", "/month").replace("day", "/day").replace("hour", "/hour"))
                except AttributeError:
                    if "FREE" in pDetail.text:
                        hText += "Please Enquire"
                    else:
                        hText += ("Price: " +
                                  pDetail.text.replace("31 days", "month").replace("29 days", "month").replace("month", "/month").replace("day", "/day").replace("hour", "/hour"))

            # Dedicated Desk
            if "Dedicated Desk" in rate.text:
                pDetail = rate.find("span", class_="price")
                sDetail = rate.find("span", class_="article_type")
                confirm[2] = True
                try:
                    pDetail2 = pDetail.find("span", class_="price_original")
                    dText += ("Price: " +
                              pDetail2.text.replace("31 days", "month").replace("29 days", "month").replace("month", "/month").replace("day", "/day").replace("hour", "/hour"))
                except AttributeError:
                    dText += ("Price: " +
                              pDetail.text.replace("31 days", "month").replace("29 days", "month").replace("month", "/month").replace("day", "/day").replace("hour", "/hour"))
            # Meeting Room
            if "Meeting" in rate.text:
                pDetail = rate.find("span", class_="price")
                sDetail = rate.find("span", class_="article_type")
                confirm[3] = True
                try:
                    pDetail2 = pDetail.find("span", class_="price_original")
                    mText += ("Price: " +
                              pDetail2.text.replace("31 days", "month").replace("29 days", "month").replace("month", "/month").replace("day", "/day").replace("hour", "/hour"))
                except AttributeError:
                    mText += ("Price: " +
                              pDetail.text.replace("31 days", "month").replace("29 days", "month").replace("month", "/month").replace("day", "/day").replace("hour", "/hour"))
                ssplit = sDetail.text.split("people")
                try:
                    mText += (", Capacity: " +
                              ssplit[0][-3:].strip(" ") + " people | ")
                except IndexError:
                    mText += (", Capacity: " + sDetail.text + " | ")

        if confirm[0] == False:
            pOffice.append("")
        else:
            pOffice.append(pText.strip(" | "))
        if confirm[1] == False:
            hDesk.append("")
        else:
            hDesk.append(hText.strip(" | "))

        if confirm[2] == False:
            dDesk.append("")
        else:
            dDesk.append(dText.strip(" | "))

        if confirm[3] == False:
            mRoom.append("")
        else:
            mRoom.append(mText.strip(" | "))

    except AttributeError:
        pOffice.append("")
        hDesk.append("")
        dDesk.append("")
        mRoom.append("")

    # Address
    try:
        addrs = soup.find("span", class_="address")
        address.append(addrs.text)
    except AttributeError:
        address.append("")

    # Facebook
    try:
        f = soup.find("div", class_="fb_web")
        fb = f.find("a")
        facebook.append(fb['href'])
    except AttributeError:
        facebook.append("")
    # Website
    try:
        w = soup.find("div", class_="web")
        web = w.find("a")
        website.append(web['href'])
    except AttributeError:
        website.append("")

    # Opening Hours
    try:
        open = soup.find("div", class_="hours")
        hour.append(open.text.replace("Fri", "Fri ").replace(
            "Sat", " Sat").replace("Sun", "Sun ").replace("Sun day", "Sunday").replace("Fri day", "Friday").replace("Monday", " Monday ").replace("Tuesday", " Tuesday ").replace("Wednesday", " Wednesday ").replace("Thursday", " Thursday ").replace("Friday", " Friday ").replace("Saturday", " Saturday ").replace("Sunday", " Sunday ").strip(" "))
    except AttributeError:
        hour.append("")

    # Amenities
    try:
        amen = soup.find("div", class_="amenities")
        amen2 = amen.find_all("div", class_="tags_group")
        aText = ""
        for am in amen2:
            h3 = am.find("h3")
            aText += h3.text + ": "
            li = am.find_all("li")
            for l in li:
                aText += l.text + ", "
            aText.strip(", ")
            aText += " | "
        amenity.append(aText.replace(",  | ", " | ").strip(" | "))
    except AttributeError:
        amenity.append("")


print(len(name))
print(len(description))
print(len(address))
print(len(location))
print(len(facebook))
print(len(website))
print(len(pOffice))
print(len(hDesk))
print(len(dDesk))
print(len(mRoom))
print(len(hour))
print(len(amenity))

# print(name)
# print(description)
# print(address)
# print(location)
# print(pOffice)
# print(facebook)
# print(website)
# print(hDesk)
# print(dDesk)
# print(mRoom)
# print(hour)
# print(amenity)

workbook = xlsxwriter.Workbook(
    'C:\\Users\\Mohy\\Desktop\\HotDesk\\ItalyCWBooking.xlsx')
worksheet = workbook.add_worksheet()

array = [name, description, address, location, facebook, website, pOffice, dDesk,
         hDesk, mRoom, hour, amenity]

row = 0

for col, data in enumerate(array):
    worksheet.write_column(row, col, data)

workbook.close()
