'''
Thats a program that alt tabs to market in game LOST ARK, click on the respective
tabs, and, for every line of product in the AH (auction house), it takes five
screenshots, one for each row: name, avg.price, recent prices, lowest prices, and cheapest.rem.

After that, the screenshots goes through a CV2 processing, and then is read with TESSERACT,
for translating these images in integers. They are then inserted in a excel file (auction.xlsx)
and reorganized in a single sheet in a second excel file (auctionresults.xlsx)


How to use:
Set your game to 1080x1024 with forced 21:9. And in the upper left side of your screen.
Put your windows bar to the left as well. Honestly, this should've been made with Client coordinates, and not screen.

Take a screenshot of every tab you want to search, and cut it with only the name
eg: Gemchest, Pets, Mounts.

This currently does not support multiple tabs. Since i  made for my own use, and i
only cared about about gathering and ship parts prices.
'''