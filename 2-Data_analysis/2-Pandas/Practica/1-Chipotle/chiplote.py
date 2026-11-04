import pandas as pd 
import numpy as np

df = pd.read_csv("C:/Users/AItor/Desktop/THE_BRIDGE_2026_DS/TEMARIO_AITOR/2-Data_analysis/2-Pandas/Practica/1-Chipotle/chiplote.csv", sep="\t")

print(df)
print(df.describe())
print(df.info())

### Step 3. Name of the max valued product
df["item_price"] = df["item_price"].str.replace("$","")
print(df.columns)
print(df["item_price"])
df["item_price"] = df["item_price"].astype(float)
print(df["item_price"].dtype)
print(df.sort_values("item_price", ascending=False))
mask= df["quantity"] == 1 
print(df[mask].sort_values(by="item_price", ascending=False))
mask=(df["item_price"] ==  11.89)
print(df[mask]["item_name"].value_counts())

### Step 4. How many products cost more than $10.00?
mask2 =(df["order_id"].nunique()) & (df["quantity"] == 1) & (df["item_price"] >=10)
print(df.loc[mask2].sort_values(by="item_price",ascending=False))


### Step 4.1: Y cuántos pedidos se han hecho con un producto de más de 10$? Es lo mismo?
mask2 = (df["quantity"] == 1) & (df["item_price"] >=10)
mask2.value_counts()


### Step 4.2: Y cuántos pedidos se han hecho de más de 20$? Es lo mismo?
mask2 = (df["item_price"] >=20)
df[mask2].count()

### Step 4.3: Y en cuántos pedidos se ha pagado más de 10$ por un mismo producto? Es lo mismo?
mask2 = (df["quantity"] == 1) & (df["item_price"] >=10)
mask2.value_counts()

### Step 5. What is the price of each item and name it unit_price. Get only item_name and unit_price
print(df[["item_name", "item_price"]])
