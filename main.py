import folium
from folium import plugins
import pandas as pd
import word_to_xlsx
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
    if j not in groups:
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






# <h2>{"Объект" if df["Объект"][i] else ""}</h2>
                                    # <p style="font-size:20px;">{df["Объект"][i] if df["Объект"][i] else ""}</p>
                                    # <img src='' style="max-width:100%">
                                    # <h2>{"Адрес" if df["Адрес"][i] else ""}</h2>
                                    # <p>{df["Адрес"][i] if df["Адрес"][i] else ""}</p>
                                    # <h2>{"ГО" if df["ГО"][i] else ""}</h2>
                                    # <p>{df["ГО"][i] if df["ГО"][i] else ""}</p>
                                    # <h2>{"Областной бюджет, тыс.рублей" if df["Областной бюджет, тыс.рублей"][i] else ""}</h2>
                                    # <p>{df["Областной бюджет, тыс.рублей"][i] if df["Областной бюджет, тыс.рублей"][i] else ""}</p>
                                    # <h2>{"Федеральный бюджет, тыс.руб." if df["Федеральный бюджет, тыс.руб."][i] else ""}</h2>
                                    # <p>{df["Федеральный бюджет, тыс.руб."][i] if df["Федеральный бюджет, тыс.руб."][i] else ""}</p>
                                    # <h2>{"Оснащение" if df["Оснащение"][i] else ""}</h2>
                                    # <p>{df["Оснащение"][i] if df["Оснащение"][i] else ""}</p>
                                    # <h2>{"Капитал. ремонт" if df["Капитал. ремонт"][i] else ""}</h2>
                                    # <p>{df["Капитал. ремонт"][i] if df["Капитал. ремонт"][i] else ""}</p>
                                    # <h2>{"Объем финансирования" if df["Объем финансирования"][i] else ""}</h2>
                                    # <p>{df["Объем финансирования"][i] if df["Объем финансирования"][i] else ""}</p>
                                    # <h2>{"Вид работ" if df["Вид работ"][i] else ""}</h2>
                                    # <p>{df["Вид работ"][i] if df["Вид работ"][i] else ""}</p>
                                    # <h2>{"Вид работ" if df["Вид работ"][i] else ""}</h2>
                                    # <p>{df["Вид работ"][i] if df["Вид работ"][i] else ""}</p>
                            #Капитальный ремонт/ строительство
print(df.columns[0])