from lib2to3.pgen2.token import COMMENT
from bs4 import BeautifulSoup
from selenium import webdriver
import threading, time, xlsxwriter, datetime

def scrollBottom(URL):
    SCROLL_PAUSE_TIME = 1
    browser = webdriver.Chrome()

    browser.get(URL)
    prevHeight = browser.execute_script("return document.body.scrollHeight")
    atBottom = False # occasionally selenium lags, this ensures that we are truly at the bottom
    while True:
        time.sleep(SCROLL_PAUSE_TIME)
        browser.execute_script("window.scrollTo(0, document.body.scrollHeight);")
        currHeight = browser.execute_script("return document.body.scrollHeight")
        if prevHeight == currHeight:
            if atBottom:
                break
            atBottom = True
        else:
            atBottom = False
        prevHeight = currHeight
    return browser

def dataInSheet(lists,sheet,position):
    listSize = len(lists)
    for x in range(listSize):
        sheet.write(x+1,position, lists[x])
    print("Finished putting in some values in excel sheet.")

def dataInSheetNum(lists,sheet,position,format):
    listSize = len(lists)
    for x in range(listSize):
        sheet.write_number(x+1,position, lists[x],format)
    print("Finished putting in all prices in excel sheet.")

def formatSheet(sheet):
    sheet.write(0,0, "Title",title_format)
    sheet.write(0,1, "Console",title_format)
    sheet.write(0,2, "Price",title_format)
    sheet.write(0,5, "Total",title_format)
    sheet.write(1,5, "{=SUM(C:C)}", money_format)    
    sheet.set_column(0,0, 45)
    sheet.set_column(1,1, 17)
    sheet.set_column(2,2, 11)
    sheet.set_column(6,6, 12)
    print("Formatted excel sheet.")

def isolateTitles(titleList):
    realTitles = []
    for x in titleList:
        realTitles.append(x.next_element.next_element.next_element)
    print("Finished gathering all title data.")
    return realTitles

def isolatePrices(priceList):
    realPrices = []
    for x in priceList:
        realPrices.append(float(x.next_element[1:]))
    print("Finished gathering all price data.")
    return realPrices

def isolateConsoles(consoleList):
    realConsoles = []
    for x in consoleList:
        realConsoles.append(str(x.next_element.next_element.next_element.next_element.next_element).strip())
    print("Finished gathering all console data.")
    return realConsoles
    

def findTitlesConsoles(bsoup):
    global titlesAndConsoles
    titlesAndConsoles = bsoup.findAll("p",attrs={"class":"title"})
    print("Found all titles.")

def findPrices(bsoup):
    global prices
    prices = bsoup.findAll("span",attrs={"class":"js-price"})
    print("Found all prices.")



#This is the Main Program Right HERE!!!!!! ------------------------------------------------------------

WEBSITEURL = input("Enter your pricecharting URL here: ")
#

#Create and names the spreadsheet and worksheet
currentDateTime = datetime.datetime.now()
#
stringDateTime = currentDateTime.strftime("/%a, %B %d, %Y at %I-%M-%S %p.xlsx")
saveLocation = input("Enter the exact location to save the file: ")
saveLocation += stringDateTime
workbook = xlsxwriter.Workbook(saveLocation)
worksheet = workbook.add_worksheet('VideoGameCollection')

#Create formats for the price of the games and the title of each header
money_format = workbook.add_format({'num_format': '$#,##0.00'})
title_format = workbook.add_format({"bold": True, "font_color": "#006eff", "font_size": 15, "align": "center", "bottom": 5, "border_color": "#006eff"})
titleCol = 0
consoleCol = 1
priceCol = 2

#parses the website and stores the information into the variable soup
soup = BeautifulSoup(scrollBottom(WEBSITEURL).page_source, 'html.parser')

t1=threading.Thread(target = findTitlesConsoles, args=(soup,))
t1.start()
t2 = threading.Thread(target = formatSheet, args=(worksheet,))
t2.start()
t3 = threading.Thread(target = findPrices, args=(soup,))
t3.start()
#isolateTitles(titlesAndConsoles)
#isolateConsoles(titlesAndConsoles)
#isolatePrices(prices)
t1.join()
t2.join()
t3.join()
t4 = threading.Thread(target = dataInSheet, args=(isolateTitles(titlesAndConsoles),worksheet,titleCol))
t4.start()
t5 = threading.Thread(target = dataInSheet, args=(isolateConsoles(titlesAndConsoles),worksheet,consoleCol))
t5.start()
t6 = threading.Thread(target = dataInSheetNum, args=(isolatePrices(prices),worksheet,priceCol,money_format))
t6.start()
t4.join()
t5.join()
t6.join()
workbook.close()

