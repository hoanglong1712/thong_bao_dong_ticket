from tkinter import *
from sys import exit
def popupError(s):
    popupRoot = Tk()
    popupRoot.after(20000, exit)
    popupButton = Button(popupRoot, text = s, font = ("Verdana", 12), bg = "yellow", command = exit)
    Label
    popupButton.pack()
    popupRoot.geometry('400x50+700+500')
    popupRoot.mainloop()

popupError('wewewe')