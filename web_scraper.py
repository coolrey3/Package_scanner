from bs4 import BeautifulSoup
from tkinter import *
import requests

root = Tk()
root.title("Fortnite News")
root.geometry("500x580")
mainTitle = Label(root, text = "Fortnite News and Stat Tracker")
mainTitle.pack()
playerName = Entry(root,width= 50)
nameLabel = Label(root,text='Enter player Name: ')
nameLabel.pack(side = "top" )
playerName.pack(side = "top")
site ='https://fortnitetracker.com'
pcPlayer= "profile/pc/"
source = requests.get(site).text
soup = BeautifulSoup(source, 'lxml')
#print(soup.prettify())
article = soup.find('article' )
print("Latest Articles")
global news
global newsbox
news = []
newsbox = Listbox(root, height=30, width=150)
newsbox.pack(side = "bottom")

class Fortnitenews(self,master):
    print("test)")


def newsresults(event):
    newsbox.delete(0, END)

    for art in soup.find_all('article'):
        time = art.time.text
        headline = art.h2.text
        link = art.h2.a['href']
        url = site + link
        news =  time + " | " + headline
        newsbox.insert(END,news)
        print("refreshed news results")

playerName.focus()
#print(time)



def searchPlayer():
    player = playerName.get()
    print(player)

root.bind("<Return>", newsresults)
searchButton = Button(root,text = "Search" ,command = searchPlayer)
searchButton.pack()
print(news)
#print(art.prettify())
title = soup.find("h2",class_="trn-article__title")
root.mainloop()
