import pandas as pd
import numpy as np
from matplotlib.pyplot import plot, show, title, xlabel, ylabel
import matplotlib.dates as mdates
from pptx import Presentation
import os


OilPrices = pd.read_csv('./CSV/crude-oil-price.csv')
OilPrices['date'] = pd.to_datetime(OilPrices['date'])

pr1 = Presentation()

pr1.save("Bogdans_Presentation.pptx")

# Open the presentation file
os.startfile("Bogdans_Presentation.pptx")

print(OilPrices.loc[0:5, ["date", "price"]])

date = OilPrices["date"]
price = OilPrices["price"]

plot(date, price)
title("oil price by time", fontsize=20)
xlabel('date', fontsize=14) 
ylabel('price', fontsize=14)

#os.startfile("presentation.pptx")

print("klar!")

show()

