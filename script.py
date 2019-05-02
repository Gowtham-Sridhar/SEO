import pandas
import sqlite3
import xlwt
from bs4 import BeautifulSoup
import re
from urllib.request import urlopen
import matplotlib.pyplot as plt
from xlrd import open_workbook


# reading excel file for input domain names
df = pandas.read_excel('SeoIn.xlsx')
# print the column names
#print (df.columns)
# get the values for a given column
urls = df['URLs'].values

# reading texts that should be ignored
fp = open("ignore.txt")
ignorelist = (((fp.read().strip()).lower()).split(" "))
# print(ignorelist)
ignoreSet = set(ignorelist)
# print(ignoreSet)

# opening database connection
conn = sqlite3.connect("myseodb.db")
cursor = conn.cursor()


# initializing database
# drop table in case exists
cursor.execute("DROP TABLE SeoData")
print("Table Existed & Deleted")

# create the table
cursor.execute(
    "CREATE TABLE SeoData (URL, Keyword text, frequency integer, density float)")
print("table created")

# exception handling for Database - only to show sample - not used everwhere
# except:
#print("DB Conn Error")
# exit()

# creating excel file to write top six words and their density
workbook = xlwt.Workbook()
kounter = 0  # used for worksheet ID

# for each entry of domain do
for i in range(len(urls)):
    url = urls[i]
    print(url)

    # validate URL
    pattern = re.compile("('http://'|'https://')")
    match = pattern.match(url)
    if match:
        print("correct URL")
    else:
        continue

    kounter = kounter + 1
    print(kounter)
    file_handle = urlopen(url)
    page = file_handle.read()
    # print(page)

    soup = BeautifulSoup(page, "html.parser")
    # print(soup)
    wordset = set()
    wordlist = []

    for script in soup(["script", "style"]):
        # text= script.extract() #remove all javascript & trailing spaces
        text = soup.get_text().lower()
        # print(text)
        allwords = (line for line in text.split())
        # print(allwords)
        chunks = (phrase.strip()
                  for line in allwords for phrase in line.split(" "))
        text = '\n'.join(chunk for chunk in chunks if chunk)
        wordlist = text.split("\n")
        #print (wordlist)

        # get all words from web page
        wordset = set(wordlist)
        # print(wordset)

        # get all words without the words to be ignored
        wordset = wordset - ignoreSet
        # print(ignoreSet)
        # print(wordset)

        # calculating total number of words in the page
        totalwords = len(wordlist)

    # creating dictionary of unique words and their frequency & density
    d = {}
    for word in wordset:
        frequency = wordlist.count(word)
        #print("word",word, "count", frequency)
        density = round(((frequency / totalwords) * 100), 3)

        # storing data to dictionary
        d[word] = density

        # create a DB record for the URL
        cursor.execute('INSERT INTO SeoData (url,keyword, frequency, density) VALUES (?,?,?,?);',
                       (url, str(word), frequency, density))
        #print("Data Inserted & Commited")

    # commit data
    conn.commit()

    # pick the top 6 high density words to store in excel
    ordered_dict_key = dict(
        sorted(d.items(), key=lambda x: x[1], reverse=True)[:6])

    # Writing topsix high density one into the excel file
    worksheet = workbook.add_sheet("Sheet" + str(kounter))

    # creating sheet headers
    row = 0
    col = 0
    worksheet.write(row, col, url)
    worksheet.write(row + 1, col, "Word")
    worksheet.write(row + 1, col + 1, "Density")
    row = 2
    for l in list(ordered_dict_key):
        col = 0
        worksheet.write(row, col, l)
        worksheet.write(row, col + 1, round(float(str(ordered_dict_key[l])), 3))

        row += 1
    workbook.save("SeoOut.xls")
    print("Written")

cursor.execute("select * from SeoData")
rows = cursor.fetchall()

for row in rows:
    print("fetched")

# closing db connection
conn.close()

# read the excel file and plot
wb = open_workbook('SeoOut.xls')
# style.use('ggplot')

for s in wb.sheets():
    print('Sheet:', s.name)
    words = []
    densitys = []

    # read title
    title = (s.cell(0, 0).value)
    x_axis_title = (s.cell(1, 0).value)
    y_axis_title = (s.cell(1, 1).value)

    # Read data from the 3rd row (as first 2 are headers)
    for row in range(2, s.nrows):
        # reads words and create list
        word = (s.cell(row, 0).value)
        words.append(word)

        # read densitys and create list
        density = (s.cell(row, 1).value)
        densitys.append(density)

    print(words, densitys)

    # plotting the data
    plt.title(title)
    plt.xlabel(x_axis_title)
    plt.ylabel(y_axis_title)
    x = range(len(words))

    for i in range(len(densitys)):
        # Plotting to our canvas
        y = densitys
        my_xticks = words
        plt.xticks(x, my_xticks)

        plt.bar(x[i], densitys[i], 1, color="blue")

    # Showing what we plotted
    plt.show()
