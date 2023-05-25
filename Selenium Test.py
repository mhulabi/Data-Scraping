from time import sleep
import xlsxwriter
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver import ActionChains
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

name = []
description = []
address = []
location = []
type = []
pOffice = []
dDesk = []
mRoom = []
hDesk = []
vOffice = []
hour = []
amenity = []

with webdriver.Chrome() as driver:
    driver.maximize_window()
    # Open URL
    # driver.get("https://www.coworker.com/search/portugal/lisbon")
    driver.get("https://www.coworker.com/search/saudi-arabia/al-khobar")
    sleep(20)

    # Setup wait for later
    wait = WebDriverWait(driver, 13)

    for j in range(1, 4):
        count = 1
        sleep(7)
        if j < 3:
            for i in range(1, 11):
                sleep(10)
                # Store the ID of the original window
                original_window = driver.current_window_handle

                # Check we don't have other windows open already
                assert len(driver.window_handles) == 1

                # Click the link which opens in a new window
                clickable = driver.find_element(
                    "xpath", '//*[@id="reactRoot"]/div[2]/div[2]/div[2]/div[%d]/div[2]' % count)
                ActionChains(driver).move_to_element(clickable).pause(
                    1).key_down(Keys.CONTROL).pause(1).click().perform()
                wait.until(EC.number_of_windows_to_be(2))

                # Loop through until we find a new window handle
                for window_handle in driver.window_handles:
                    if window_handle != original_window:
                        driver.switch_to.window(window_handle)
                        break

                # Wait for the new tab to finish loading content
                print("here")
                sleep(10)
                html = driver.page_source
                # print(html.encode("utf-8"))
                soup = BeautifulSoup(html, "html.parser")
                # soup.encode("utf-8")

                # Name
                try:
                    names = soup.find("h1", id="old_header")
                    # .encode("utf-8"))
                    name.append(names.text.strip("\n").strip())
                except AttributeError:
                    name.append("")

                # Description
                try:
                    descs = soup.find(
                        "div", class_="space-full-description-inner")
                # Encoding required from print
                    description.append(descs.text.strip("\n").strip(" ").replace(
                        "\n", " ").replace("  ", " "))  # .encode("utf-8"))
                except AttributeError:
                    description.append("")
                # print(description)

                # Address
                try:
                    addrs = soup.find(
                        "div", class_="col-12 pad-none space-address-con")
                # print(addrs)
                    address.append(addrs.text.replace("NoView on Map", "").replace("View on Map", "").replace(
                        "No View on\nMap", "").replace("No View onMap", "").replace(
                        "View on\nMap", "").strip(" "))
                except AttributeError:
                    address.append("")
                # print(address)

                # Location
                try:
                    loc = soup.find("title")
                    split = loc.text.split(',')
                    location.append(
                        split[-1].replace(" | Coworker", "").strip(" "))
                except AttributeError:
                    location.append("")

                # Rates
                rates = soup.find_all(
                    "div", class_="col-12 pad-none membership-fees-mobile-outer")
                confirm = [False, False, False, False]
                for rate in rates:
                    # Private Office
                    if "Private" in rate.text:
                        confirm[0] = True
                        pOffice.append(rate.text.replace("\nMembership: Private Offices\n", "").replace("Price", "Price: ").replace("\n\n#People\n", " Persons: ").replace("\n", "").replace("Monthly", "Month").replace(
                            "Month", " | Month").replace("Year", " | Year").replace("Week", " | Week").replace("Day", " | Day").replace("Please enquire for best available pricing", "Please Enquire").strip(" | "))  # .strip(" |"))
                    # Hot Desk
                    elif "Hot" in rate.text:
                        confirm[1] = True
                        hDesk.append(rate.text.replace("\nMembership: Hot Desks\n", "").replace("Price", "Price: ").replace("\n\n#People\n", " Persons: ").replace("\n", "").replace("Monthly", "Month").replace(
                            "Month", " | Month").replace("Year", " | Year").replace("Week", " | Week").replace("Day", " | Day").replace("Please enquire for best available pricing", "Please Enquire").strip(" | "))  # .replace("lease enquire for best available pricing", "Please Enquire").replace("onth", "Month").replace("ay", "Day"))
                    # Dedicated Desk
                    if "Dedicated" in rate.text:
                        confirm[2] = True
                        dDesk.append(rate.text.replace("\nMembership: Dedicated Desks\n", "").replace("Price", "Price: ").replace("\n\n#People\n", " Persons: ").replace("\n", "").replace("Monthly", "Month").replace(
                            "Month", " | Month").replace("Year", " | Year").replace("Week", " | Week").replace("Day", " | Day").replace("Please enquire for best available pricing", "Please Enquire").strip(" | "))  # .strip(" |"))
                    # Virtual Office
                    if "Virtual" in rate.text:
                        confirm[3] = True
                        vOffice.append(rate.text.replace("\nMembership: Virtual Office\n", "").replace("Price", "Price: ").replace("\n\n#People\n", " Persons: ").replace("\n", "").replace("Monthly", "Month").replace(
                            "Month", " | Month").replace("Year", " | Year").replace("Week", " | Week").replace("Day", " | Day").replace("Please enquire for best available pricing", "Please Enquire").strip(" | "))  # .strip(" |"))

                if confirm[0] == False:
                    pOffice.append("")
                if confirm[1] == False:
                    hDesk.append("")
                if confirm[2] == False:
                    dDesk.append("")
                if confirm[3] == False:
                    vOffice.append("")

                # Meeting Room
                try:
                    rooms = soup.find("div", id="meeting-rooms")
                    test = rooms.text
                    mRoom.append("Please Enquire")
                except AttributeError:
                    mRoom.append("")

                # Hour
                try:
                    times = soup.find(
                        "div", class_="col-12 pad-none space-member-content")
                    hour.append(times.text.strip("\n").strip(" Opening Hours").replace(
                        "Sat", " Sat").replace("Sun", " Sun").replace("  Sun", " Sun"))
                except AttributeError:
                    hour.append("")

                # Amenity
                try:
                    totalAms = ""
                    ams = soup.find_all(
                        "div", class_="col-12 pad-none space-amenities-outer")
                    for am in ams:
                        totalAms += am.text
                    amenity.append(totalAms.replace(
                        "\n", "").replace("Classic Basics", "Classic Basics: ").replace("Seating", "| Seating: ").replace("Relax Zones", "| Relax Zones: ").replace("Facilities", "| Facilities: ").replace("Catering", "| Catering: ").replace("Caffeine Fix", "| Caffeine Fix: ").replace("Community", "| Community: ").replace("Equipment", "| Equipment: ").replace("Cool Stuff", "| Cool Stuff: ").replace("Transportation", "| Transportation: ").replace("Accessibility", "| Accessibility: ").replace("Wheelchair | Accessibility: ", "Wheelchair Accessibility").strip("| "))
                except AttributeError:
                    amenity.append("")
                driver.close()

                # Switch back to the old tab or window
                driver.switch_to.window(original_window)
                print(count)
                sleep(2)
                count += 1
        else:
            for i in range(1, 11):
                sleep(10)
                # Store the ID of the original window
                original_window = driver.current_window_handle

                # Check we don't have other windows open already
                assert len(driver.window_handles) == 1

                # Click the link which opens in a new window
                clickable = driver.find_element(
                    "xpath", '//*[@id="reactRoot"]/div[2]/div[2]/div[2]/div[%d]/div[2]' % count)
                ActionChains(driver).move_to_element(clickable).pause(
                    1).key_down(Keys.CONTROL).pause(1).click().perform()
                wait.until(EC.number_of_windows_to_be(2))

                # Loop through until we find a new window handle
                for window_handle in driver.window_handles:
                    if window_handle != original_window:
                        driver.switch_to.window(window_handle)
                        break

                # Wait for the new tab to finish loading content
                print("here")
                sleep(15)
                html = driver.page_source
                # print(html.encode("utf-8"))
                soup = BeautifulSoup(html, "html.parser")
                # soup.encode("utf-8")

                # Name
                try:
                    names = soup.find("h1", id="old_header")
                    # .encode("utf-8"))
                    name.append(names.text.strip("\n").strip())
                except AttributeError:
                    name.append("")

                # Location
                try:
                    loc = soup.find("title")
                    split = loc.text.split(',')
                    location.append(
                        split[-1].replace(" | Coworker", "").strip(" "))
                except AttributeError:
                    location.append("")

                # Description
                try:
                    descs = soup.find(
                        "div", class_="space-full-description-inner")
                # Encoding required from print
                    description.append(descs.text.strip("\n").strip(" ").replace(
                        "\n", " ").replace("  ", " "))  # .encode("utf-8"))
                except AttributeError:
                    description.append("")
                # print(description)

                # Address
                try:
                    addrs = soup.find(
                        "div", class_="col-12 pad-none space-address-con")
                # print(addrs)
                    address.append(addrs.text.replace("NoView on Map", "").replace("View on Map", "").replace(
                        "No View on\nMap", "").replace("No View onMap", "").replace(
                        "View on\nMap", "").strip(" "))
                except AttributeError:
                    address.append("")
                # print(address)

                # Rates
                rates = soup.find_all(
                    "div", class_="col-12 pad-none membership-fees-mobile-outer")
                confirm = [False, False, False, False]
                for rate in rates:
                    # Private Office
                    if "Private" in rate.text:
                        confirm[0] = True
                        pOffice.append(rate.text.replace("\nMembership: Private Offices\n", "").replace("Price", "Price: ").replace("\n\n#People\n", " Persons: ").replace("\n", "").replace("Monthly", "Month").replace(
                            "Month", " | Month").replace("Year", " | Year").replace("Week", " | Week").replace("Day", " | Day").replace("Please enquire for best available pricing", "Please Enquire").strip(" | "))  # .strip(" |"))
                    # Hot Desk
                    elif "Hot" in rate.text:
                        confirm[1] = True
                        hDesk.append(rate.text.replace("\nMembership: Hot Desks\n", "").replace("Price", "Price: ").replace("\n\n#People\n", " Persons: ").replace("\n", "").replace("Monthly", "Month").replace(
                            "Month", " | Month").replace("Year", " | Year").replace("Week", " | Week").replace("Day", " | Day").replace("Please enquire for best available pricing", "Please Enquire").strip(" | "))  # .replace("lease enquire for best available pricing", "Please Enquire").replace("onth", "Month").replace("ay", "Day"))
                    # Dedicated Desk
                    if "Dedicated" in rate.text:
                        confirm[2] = True
                        dDesk.append(rate.text.replace("\nMembership: Dedicated Desks\n", "").replace("Price", "Price: ").replace("\n\n#People\n", " Persons: ").replace("\n", "").replace("Monthly", "Month").replace(
                            "Month", " | Month").replace("Year", " | Year").replace("Week", " | Week").replace("Day", " | Day").replace("Please enquire for best available pricing", "Please Enquire").strip(" | "))  # .strip(" |"))
                    # Virtual Office
                    if "Virtual" in rate.text:
                        confirm[3] = True
                        vOffice.append(rate.text.replace("\nMembership: Virtual Office\n", "").replace("Price", "Price: ").replace("\n\n#People\n", " Persons: ").replace("\n", "").replace("Monthly", "Month").replace(
                            "Month", " | Month").replace("Year", " | Year").replace("Week", " | Week").replace("Day", " | Day").replace("Please enquire for best available pricing", "Please Enquire").strip(" | "))  # .strip(" |"))

                if confirm[0] == False:
                    pOffice.append("")
                if confirm[1] == False:
                    hDesk.append("")
                if confirm[2] == False:
                    dDesk.append("")
                if confirm[3] == False:
                    vOffice.append("")

                # Meeting Room
                try:
                    rooms = soup.find("div", id="meeting-rooms")
                    test = rooms.text
                    mRoom.append("Please Enquire")
                except AttributeError:
                    mRoom.append("")

                # Hour
                try:
                    times = soup.find(
                        "div", class_="col-12 pad-none space-member-content")
                    hour.append(times.text.strip("\n").strip(" Opening Hours").replace(
                        "Sat", " Sat").replace("Sun", " Sun").replace("  Sun", " Sun"))
                except AttributeError:
                    hour.append("")

                # Amenity
                try:
                    totalAms = ""
                    ams = soup.find_all(
                        "div", class_="col-12 pad-none space-amenities-outer")
                    for am in ams:
                        totalAms += am.text
                    amenity.append(totalAms.replace(
                        "\n", "").replace("Classic Basics", "Classic Basics: ").replace("Seating", "| Seating: ").replace("Relax Zones", "| Relax Zones: ").replace("Facilities", "| Facilities: ").replace("Catering", "| Catering: ").replace("Caffeine Fix", "| Caffeine Fix: ").replace("Community", "| Community: ").replace("Equipment", "| Equipment: ").replace("Cool Stuff", "| Cool Stuff: ").replace("Transportation", "| Transportation: ").replace("Accessibility", "| Accessibility: ").replace("Wheelchair | Accessibility: ", "Wheelchair Accessibility").strip("| "))
                except AttributeError:
                    amenity.append("")
                driver.close()

                # Switch back to the old tab or window
                driver.switch_to.window(original_window)
                print(count)
                sleep(2)
                count += 1
        # Clicking next button
        if j == 1:
            clickable2 = driver.find_element(
                "xpath", '//*[@id="reactRoot"]/div[2]/nav/ul/li[4]/button')
            ActionChains(driver).move_to_element(clickable2).pause(
                1).click().perform()
        elif j > 1 and j < 3:
            clickable2 = driver.find_element(
                "xpath", '//*[@id="reactRoot"]/div[2]/nav/ul/li[5]/button')
            ActionChains(driver).move_to_element(clickable2).pause(
                1).click().perform()


#Write in Excel
workbook = xlsxwriter.Workbook(
    'C:\\Users\\Mohy\\Desktop\\HotDesk\\SaudiCoworker3.xlsx')
worksheet = workbook.add_worksheet()

array = [name, description, address, location, pOffice, dDesk,
         hDesk, vOffice, mRoom, hour, amenity]

row = 0

for col, data in enumerate(array):
    worksheet.write_column(row, col, data)

workbook.close()
