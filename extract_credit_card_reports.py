from selenium import webdriver
from selenium.webdriver.common.action_chains import ActionChains
from selenium.webdriver.common.keys import Keys
from time import sleep
import os
import shutil


DOWNLOADS_FOLDER = "D:\\Users\\danid\\Downloads"  # browser downloads folder
DESTINATION_FOLDER = "Cards"  # folder to move all your data (repeated filenames will be overwritten)


cards_to_skip = []  # [1, 2]

def start_browser():
    # feel free to change this to your favorite browser (only tested in Chrome)
    return webdriver.Chrome()


def ctrl_click(driver, element):
    ActionChains(driver).key_down(Keys.CONTROL).click(element).key_up(Keys.CONTROL).perform()


if __name__ == "__main__":
    # make sure downloads folder exists
    if not os.path.isdir(DOWNLOADS_FOLDER):
        print("Downloads folder '{}' does not exist.".format(DESTINATION_FOLDER))
        print("Make sure to select the right downloads folder inside the code.")
        raise Exception("Downloads folder not found")

    # create destination folder if necessary
    if not os.path.isdir(DESTINATION_FOLDER):
        print("Creating folder: '{}'".format(DESTINATION_FOLDER))
        os.mkdir(DESTINATION_FOLDER)

    # open browser
    driver = start_browser()
    driver.maximize_window()

    # open BoA website
    driver.get("https://www.bankofamerica.com")

    # wait for manual entry of user and password
    while True:
        print("\nLOG INTO YOUR ACCOUNT SO THAT YOUR CREDIT CARD REPORTS CAN BE DOWNLOADED...")
        print("Press ENTER when you are done...")
        resp = input()
        if driver.current_url.startswith("https://secure.bankofamerica.com/myaccounts/signin/signIn.go"):
            break
        else:
            print("Login failed... Try again!")

    # get card files and save them in folder structure
    card_names_set = set()
    while True:
        # click on card
        cards = driver.find_element_by_class_name("AccountItems").find_elements_by_tag_name("li")
        new_card_found = False
        for card in cards:
            a = card.find_element_by_class_name("AccountItem").find_element_by_class_name("AccountName").find_element_by_tag_name("a")
            card_name = a.get_attribute("innerHTML").strip()
            if card_name not in card_names_set:
                card_names_set.add(card_name)
                new_card_found = True
                destination_folder = os.path.join(DESTINATION_FOLDER, card_name)
                if not os.path.isdir(destination_folder):
                    os.mkdir(destination_folder)
                break
        if not new_card_found:  # parsed all cards
            break
        print("Downloading data of card {}. {}...".format(len(card_names_set), card_name))
        if len(card_names_set) in cards_to_skip:
            print("Skipping...")
            continue
        a.click()
        sleep(2)

        # click on download button
        try:
            # debit card
            debit = True
            download_link = driver\
                .find_element_by_class_name("transaction-links")\
                .find_element_by_class_name("right-links")\
                .find_element_by_class_name("download-links")\
                .find_element_by_tag_name("a")
            driver.execute_script("arguments[0].scrollIntoView();", download_link)
            download_link.click()  # we only click once
        except Exception:
            # credit card
            debit = False
            download_link = driver\
                .find_element_by_class_name("print-n-export-links")\
                .find_element_by_class_name("download-width")\
                .find_element_by_tag_name("a")
            driver.execute_script("arguments[0].scrollIntoView();", download_link)
            download_link.click()  # we click once at the beginning and will click after every download

        # choose what partitions to download
        menu = driver.find_element_by_class_name("icon-legend-download")
        if debit:
            # debit card
            menu.find_element_by_name("download_file_in_this_format_CSV").click()  # choose download as CSV
            options = menu.find_element_by_id("select_txnperiod").find_elements_by_tag_name("option")
        else:
            # credit card
            menu.find_element_by_name("download_file_in_this_format_COMMA_DELIMITED").click()  # choose download as CSV
            options = menu.find_element_by_id("select_transaction").find_elements_by_tag_name("option")
        # download all files
        for option in options:
            option.click()  # choose date range
            option_name = option.get_attribute("innerHTML").replace("\n", "").replace("\t", "").replace("\\", "-").replace("/", "-").strip()
            menu.find_element_by_class_name("submit-download").click()  # download report

            sleep(1)
            filename = max([DOWNLOADS_FOLDER + "\\" + f for f in os.listdir(DOWNLOADS_FOLDER) if not f.endswith(".tmp")], key=os.path.getctime)
            destination = os.path.join(destination_folder, option_name + ".csv")
            print(filename, "\t\t->\t\t", destination)
            retries = 0
            max_retries = 10
            copied = False
            while retries < max_retries:
                try:
                    if retries > 0:
                        print("  File downloading, retrying...")
                    retries += 1
                    shutil.move(filename, destination)
                    print("  Done!")
                    copied = True
                    break
                except Exception:
                    sleep(1)
            if not copied:
                print("  Error!")
            
            if not debit:
                # credit card
                download_link.click()  # after downloading dialog goes away

        # go back and repeat for all cards
        driver.back()

    print("\nAll available data copied to folder '{}'".format(DESTINATION_FOLDER))