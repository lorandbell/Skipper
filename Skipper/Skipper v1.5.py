from lxml import html
import requests, sys, string, time, os, warnings, csv
warnings.filterwarnings("ignore")

print ("Select which file number you wish to open") 
i=0
for file in os.listdir("./"):
    print str(i) +". " + (file)
    i += 1
fileNo = raw_input('>')


print "Opening " + os.listdir("./")[int(fileNo)]
'''
csv = open(os.listdir("./")[int(fileNo)], 'r')

DB,db_buffer =[],[]
for row in csv:
    db_buffer.append('' if row.split(",")[0] == None else row.split(",")[0])
    db_buffer.append('' if row.split(",")[1] == None else row.split(",")[1][:-1])
    DB.append(db_buffer)
    db_buffer = []
del DB[0]'''

fileReader = csv.reader(open(os.listdir("./")[int(fileNo)]), delimiter=",")
fileReader.next()
DB,db_buffer =[],[]
for Name, City in fileReader:
    DB.append( [ Name , City ] )



def GetPhone( Name, City):  

    NameString = string.replace( Name , " ", "+" )
    CityString = string.replace( City , " ", "+" )
    CityString = string.replace( CityString , ",", "%2C" )

    SearchPage = 1
    searchUrl = "http://www.canada411.ca/search/si/%s/%s/%s/?pgLen=25" % (SearchPage, NameString, CityString )
    page = requests.get(searchUrl)
    tree = html.fromstring(page.content)
    Array = []

    #Determine how if results are found
    try:
        # see how many names match
        NoNameMacthes = tree.xpath('//*[@id="c411Body"]/div[2]/div[1]/div[3]/div[1]/h1/text()')[0].split(" ")[0]
        print NoNameMacthes + " contacts found for " + Name

    except:
        try:
            print "Contact Found : " +tree.xpath('//*[@id="contact"]/h1/text()')[0]
            NoNameMacthes = 1
        except:
            print  "No one found for %s %s" % ( Name, " in %s" % City if(City != "") else "")
            NoNameMacthes = 0
            raw_input()

    if NoNameMacthes == 1:
        ph = ''.join(tree.xpath('//*[@id="contact"]/div[2]/div/ul/li[1]/span/text()'))
        ct = ''.join(tree.xpath('//*[@id="contact"]/h1/text()'))
        ad = ''.join(tree.xpath('//*[@id="contact"]/div[1]/text()'))
        Array.append([ct,ph,ad])

    NoNameMacthes = int(NoNameMacthes)
    pageCount = 1
    while (SearchPage-1)*25 < NoNameMacthes & NoNameMacthes != 1 | 0:
       
        searchUrl = "http://www.canada411.ca/search/si/%s/%s/%s/?pgLen=25" % (SearchPage, NameString, CityString )
        page = requests.get(searchUrl)
        tree = html.fromstring(page.content)

        while pageCount<=25:
                if (SearchPage-1)*25 + pageCount > NoNameMacthes: break

                ph = ''.join(tree.xpath('//*[@id="ContactPhone%s"]/text()'%pageCount))
                ct = ''.join(tree.xpath('//*[@id="ContactName%s"]/a/text()'%pageCount))
                ad = ''.join(tree.xpath('//*[@id="ContactAddress%s"]/text()'%pageCount))
                street= ad.split(ad.split(" ")[-4])[0]
                city = ad.split(" ")[-4]
                pr = ad.split(" ")[-3]
                postal= ad.split(ad.split(" ")[-4])[1]
                Array.append([ct,ph,street,city,pr,postal])
                # print "Extracting information " + str((SearchPage-1)*25+ pageCount) +" of " + str(NoNameMacthes)
                pageCount +=1
        pageCount = 1
        SearchPage += 1
    print str(NoNameMacthes) + " contacts extracted"
    return Array


Array = []

for i in DB:
    for i in GetPhone( i[0],i[1]):
        Array.append(i)


print "Saving excel file"
import xlwt
book = xlwt.Workbook(encoding="utf-8")
sheet1 = book.add_sheet("Sheet 1")

row = 0
for l in Array:
    col = 0
    for e in l:
        if e:
          sheet1.write(row, col, e)
        col+=1
    row+=1
book.save("Name Search.xls")
raw_input("File Saved")
