import pandas as pd
import numpy as np
from matplotlib.pyplot import plot, show, title, xlabel, ylabel, savefig
import matplotlib.dates as mdates
from pptx import Presentation
import os


OilPrices = pd.read_csv('./CSV/crude-oil-price.csv')
OilPrices['date'] = pd.to_datetime(OilPrices['date'])

print(OilPrices.loc[0:5, ["date", "price"]])

date = OilPrices["date"]
price = OilPrices["price"]

# Get highest and lowest prices
highest_price = price.max()
lowest_price = price.min()
highest_date = date[price.idxmax()]
lowest_date = date[price.idxmin()]

print(f"Highest price: ${highest_price} on {highest_date}")
print(f"Lowest price: ${lowest_price} on {lowest_date}")

plot(date, price)
title("date by price", fontsize=20)
xlabel('date', fontsize=14) 
ylabel('price', fontsize=14)

# Save the plot as an image
savefig("oil_price_graph.png", dpi=100, bbox_inches='tight')

# Calculate daily percentage change
percentage_change = price.pct_change() * 100

# Create percentage change graph
plot(date, percentage_change)
title("Daily Percentage Change in Oil Price", fontsize=20)
xlabel('date', fontsize=14)
ylabel('percentage change (%)', fontsize=14)

# Save the percentage change plot
savefig("oil_price_percentage_change.png", dpi=100, bbox_inches='tight')

# Create presentation and add the graph
pr1 = Presentation()
slide = pr1.slides.add_slide(pr1.slide_layouts[5])
slide2 = pr1.slides.add_slide(pr1.slide_layouts[5])
slide3 = pr1.slides.add_slide(pr1.slide_layouts[5])
slide4 = pr1.slides.add_slide(pr1.slide_layouts[5])

# Add first graph
left = top = 0
pic = slide2.shapes.add_picture("oil_price_graph.png", left, top, width=9144000, height=6858000)

# Add percentage change graph to third slide
pic2 = slide4.shapes.add_picture("oil_price_percentage_change.png", left, top, width=9144000, height=6858000)

# Add title to second slide
title = slide.shapes.title
title.text = "Oil Price Over Time"
title2 = slide2.shapes.title
title2.text = "Highest and lowest price"

pr1.save("Bogdans_Presentation.pptx")

# Open the presentation file
os.startfile("Bogdans_Presentation.pptx")

print("klar!")

show()

