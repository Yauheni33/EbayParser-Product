import requests
import xlsxwriter
from bs4 import BeautifulSoup
import openpyxl
import json
from ebaysdk.shopping import Connection as Shopping

header = {
    'User-Agent': 'Mozilla/5.0 (Macintosh; Intel Mac OS X 10_13_5) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/11.1.1 Safari/605.1.15',
    'Accept-Language': 'en-us',
}

head = dict(header)
url = 'https://www.ebay.com/sch/i.html?_from=R40&_sacat=0&_nkw=door+stopper&LH_BIN=1&_fsrp=1&LH_PrefLoc=3&_ipg=200'
api = "http://open.api.ebay.com/shopping?callname=GetSingleItem&responseencoding=JSON&appid=Benjamin-TrendWat-PRD-f2466ad44-bc17cfa6&siteid=0&version=981&IncludeSelector=Compatibility,Description,Details,ItemSpecifics,TextDescription,HighBidder.FeedbackPrivate,HighBidder.FeedbackScore&ItemID="

page = requests.get(url, headers=head)
html = BeautifulSoup(page.text, "html.parser")
print(page)

allItem = html.findAll("li", {"class": "s-item"})

row = 4

workbook = xlsxwriter.Workbook("EbayProduct.xlsx")
worksheet = workbook.add_worksheet()

for page in range(4):
    print("PAGE: ", page)
    allItem = html.findAll("li", {"class": "s-item"})
    for i in allItem:
        print("Товар: ", (row - 3))
        print(i.find("a", {"class": "s-item__link"})['href'])
        feedProduct = "officeproducts"
        newPage = BeautifulSoup((requests.get(i.find("a", {"class": "s-item__link"})['href'])).text, "html.parser")
        IDproduct = newPage.find("div", {"id": "descItemNumber"}).text

        apiPage = requests.get(api + IDproduct)
        apiPage = json.loads(apiPage.content)
        image = apiPage['Item']['PictureURL'][0]
        description = apiPage['Item']['Description']
        name = apiPage['Item']['Title']
        try:
            for itemSpecifics in apiPage['Item']['ItemSpecifics']['NameValueList']:
                if itemSpecifics['Name'] == 'MPN':
                    mpn = itemSpecifics['Value']
                if itemSpecifics['Name'] == 'Brand':
                    brand = itemSpecifics['Value']
        except:
            mpn = ['', '']
            brand = ['', '']
            print("Нет данных")

        with open("file.html", "w") as file:
            file.write(str(newPage))

        manufacturer = brand[0]
        itemType = "door-stops"
        x = 0
        try:
            x = float(newPage.find("span", {"id": "fshippingCost"}).text.replace('US $', '').replace(',', '.').replace('C $', ''))
        except Exception as Error:
            print("Бесплатная доставка: ", Error)

        f = newPage.find("span", {"class": "notranslate"}).text.replace('US $', '').replace(',', '.').replace('GBP ',
                                                                                                              '').replace(
            'C ', '')
        f = f.replace('$', '')

        price = 2 * (float(f) + x)
        update = "Update"
        ID = 0
        #.text.replace('US $', '').replace(',', '.').replace('GBP ', '').replace('C ', '')
        try:
            ID = newPage.find("h2", {"itemprop": "gtin13"}).text
        except Exception as Error:
            print("Нет ID")
        '''
        name = newPage.find("h1", {"class": "it-ttl"}).text
        brand = ""
        try:
            brand = newPage.find("h2", {"itemprop": "brand"}).find("span").text
        except:
            print("Нет бренда")
        mpn = ""
        try:
            mpn = newPage.find("h2", {"itemprop": "mpn"}).text
        except:
            print("Нет MPN")
        fullDescription = ""
        try:
            descriptionPage = BeautifulSoup((requests.get(newPage.find("iframe", {"sandbox": "allow-scripts allow-popups allow-popups-to-escape-sandbox allow-same-origin"})['src'])).text, "html.parser")
        except:
            descriptionPage = BeautifulSoup((requests.get(newPage.find("a", {
                "class": "btn btn-ter u-padT10 "})['href'])).text,
                                            "html.parser")
        for description in descriptionPage.findAll(attrs={"id": "description"}):
            fullDescription += description.text
        print(feedProduct)
        print(brand)
        print(name)
        print(manufacturer)
        print(mpn)
        print(itemType)
        print(price)
        print(image)
        print(ID)
        print(fullDescription)
        '''
        worksheet.write('A' + str(row), feedProduct)
        worksheet.write('B' + str(row), (row - 3))
        worksheet.write('C' + str(row), brand[0])
        worksheet.write('D' + str(row), name)
        worksheet.write('E' + str(row), manufacturer)
        worksheet.write('F' + str(row), mpn[0])
        worksheet.write('G' + str(row), itemType)
        worksheet.write('H' + str(row), price)
        worksheet.write('L' + str(row), image)
        worksheet.write('M' + str(row), ID)
        worksheet.write('N' + str(row), description)
        row += 1
    try:
        url = html.findAll(attrs={"class": "x-pagination__control"})[1]['href']
        page = requests.get(url, headers=head)
        html = BeautifulSoup(page.text, "html.parser")
    except:
        print("Ошибка в странице")
        exit()