#! python3

from tkinter import *
import tkinter.messagebox as MB
from bs4 import BeautifulSoup
import requests, sys, webbrowser, openpyxl, subprocess, re

names = []
prices = []
ratings = []
links = []

def helloCallBack():
    stuff = e1.get()
    MB.showinfo(title=None, message = 'Searching...')
    search(stuff)
    makeSpreadsheet(stuff)

def findRegex(text):
    ratingRegex = re.compile(r'\d\.\d')
    rat = ratingRegex.search(text)
    return rat.group()


def search(text):
    res = requests.get('http://www.amazon.com/s/keywords=' + text)
    res.raise_for_status()

    # Retrieve top search result links.
    soup = BeautifulSoup(res.text)

    # Open a browser tab for each result
    linkElems = soup.select('.s-access-detail-page')
    numOpen = len(linkElems)
    #numOpen = min(10, len(linkElems))
    for i in range(numOpen):
        links.append(linkElems[i].get('href'))
        r = requests.get(linkElems[i].get('href'))
        r.raise_for_status()
        s = BeautifulSoup(r.text)
        price = s.select('#priceblock_ourprice')
        if price:
            prices.append(price[0].text)
        else:
            prices.append(" ")
        name = s.select('#productTitle')
        names.append(name[0].text)
        rating = s.select('.a-popover-trigger.a-declarative span')
        finalRat = findRegex(rating[0].text)
        ratings.append(finalRat)


def makeSpreadsheet(text):
    wb = openpyxl.Workbook()
    sheet = wb.get_active_sheet()
    sheet.title = text

    sheet['A1'] = 'Product Name'
    sheet['B1'] = 'Price'
    sheet['C1'] = 'Rating'
    sheet['D1'] = 'Link'
    
    j = 2
    while((j-2) < len(names)):
        sheet['A' + str(j)] = names[j-2]
        sheet['B' + str(j)] = prices[j-2]
        sheet['C' + str(j)] = ratings[j-2]
        sheet['D' + str(j)] = links[j-2]
        j = j + 1       

    filename = (text.replace(' ', '.') + ('.results.xlsx'))

    wb.save(filename)

    subprocess.Popen(['start', filename], shell=True)

    names[:] = []
    prices[:] = []
    ratings[:] = []
    links[:] = []
            

#creates window
root = Tk()

# sets root window
root.title("Amazon Scraper")
root.geometry('400x75')

app = Frame(root)
app.grid()

l1 = Label(app, text = 'Enter search term:')
l1.grid(column=0, row=0, pady=20, padx=7)
e1 = Entry(app, bd = 5)
e1.grid(column=1, row=0, pady=20, padx=7)

b1 = Button(app, text = 'Search', command = helloCallBack)
b1.grid(column=2, row=0, pady=20, padx=7)

# starts event loop
root.mainloop()




