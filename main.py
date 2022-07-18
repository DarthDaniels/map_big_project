import folium
import pandas as pd
import os
from tkinter import Tk

df = pd.read_excel('table.xlsx', sheet_name='Sheet1')
df = df.fillna(0)
WIDTH = Tk().winfo_screenwidth()
COLORS = ['orange', 'green', 'darkpurple', 'gray', 'blue', 'purple', 'black', 'darkblue',
          'cadetblue', 'lightblue', 'lightgreen', 'beige', 'lightgray',
          'darkgreen',  'lightred', 'red', 'white', 'pink', 'darkred']

groups = dict()
layers = dict()

for j in df["Тип объекта"]:
    if j not in groups and j != 0:
        groups[j] = len(groups)
        layers[j] = (folium.FeatureGroup(name=j))

print(groups)

sakh_map = folium.Map(
    location=[50.3834403, 144.026794],
    zoom_start=6,
)

Layers = folium.LayerControl()

for i in range(len(df["Долгота"])):
    if df["Долгота"][i] != 0:
        d = str()
        for j in range(1, len(df.columns)):
            if df[df.columns[j]][i] != 0 and df.columns[j] != "Долгота" and df.columns[j] != "Широта":
                d += f"<h3> {df.columns[j]} </h3>\n"
                d += f"<h4> {df[df.columns[j]][i]} </h4>\n"
        marker = folium.Marker([float(df["Долгота"][i]), float(df["Широта"][i])],
                               icon=folium.Icon(color=COLORS[groups[df["Тип объекта"][i]]]),
                               popup=(f"""
                               <head>
                                    <style>
                                     .tt{{
                                       overflow-y: scroll !important;
                                       max-height: {WIDTH//3}px;
                                     }} 
                                     </style>
                                </head>
                               <body>
                               <div class='tt', style="width: {WIDTH//4}px;">
                               {d}
                               </div>
                               </body>
                            """))
        marker.add_to(layers[df["Тип объекта"][i]])

for layer in layers:
    sakh_map.add_child(layers[layer])

sakh_map.add_child(folium.map.LayerControl())
sakh_map.save("map1.html")
os.system("C:/Users/a.reukov/Desktop/Данил/build_map/map1.html")
