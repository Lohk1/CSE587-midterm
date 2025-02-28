import pandas as pd
import re
import os

file_path = r"C:\Users\25073\OneDrive\Desktop\CSE 587\midterm\mr-porter.csv"
df = pd.read_csv(file_path, encoding="utf-8", low_memory=False)

df_cleaned = df.dropna(subset=["brand", "description"])
df_cleaned["description"] = df_cleaned["description"].apply(lambda x: re.sub(r"[^a-zA-Z0-9\s]", "", str(x)))
df_cleaned["description"] = df_cleaned["description"].str.lower()

brand_counts = df_cleaned["brand"].value_counts().reset_index()
brand_counts.columns = ["Brand", "Count"]

output_path = os.path.join(os.path.dirname(file_path), "mr-porter_brand_counts.xlsx")
brand_counts.to_excel(output_path, index=False)

selected_brands = [
    "GUCCI", "SAINT LAURENT", "POLO RALPH LAUREN", "MONCLER", "CELINE HOMME",
    "BALENCIAGA", "BOTTEGA VENETA", "NIKE", "LOEWE", "BURBERRY", "GIVENCHY",
    "LULULEMON", "ALEXANDER MCQUEEN", "A.P.C.", "CANADA GOOSE", "STONE ISLAND",
    "MONTBLANC", "DUNHILL"
]

df_filtered = df[df["brand"].isin(selected_brands)]

output_path = os.path.join(os.path.dirname(file_path), "mr-porter_filtered.xlsx")
df_filtered.to_excel(output_path, index=False)
