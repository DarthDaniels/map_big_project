def clear(string: str, trash: str) -> str:
    letter1 = 0
    if string.__contains__(trash):
        while letter1 != len(trash):
            string = string[0:string.rfind(trash[letter1])] + string[string.rfind(trash[letter1]) + 1:]
            letter1 += 1
    return string


def word_to_xlsx():

    from docx import Document
    import re
    import dadata as dt
    import folium
    import pandas as pd
    import openpyxl

    document = Document('ОБЪЕКТЫ САХАЛИНСКОЙ ОБЛАСТИ.docx')
    hand_data_txt = open("hand_data.txt", "r", encoding="utf-8")

    multi_space_pattern = re.compile(r'\s{2,}')
    big_data = []

    cur = -1
    keys = []

    for table in document.tables:
        columns = []
        for row in table.rows:
            data = [multi_space_pattern.sub(' ', i.text.strip()) for i in row.cells]
            if data.__contains__('№'):
                keys = data
            elif data[1] != '' and not data.__contains__("Оснащение") and \
                    not (data[2][0:5].__contains__("Капит") and not data[2].__contains__("М")
                            and not data[2].__contains__("Д")
                            and not data[1].__contains__("Капит")):
                new_dict = {}
                for elem in range(1, len(keys)):
                    new_dict[keys[elem]] = clear(clear(data[elem], "\n"), "\xa0")
                big_data.append(new_dict)
            elif data[2][0:5].__contains__("Капит") and not data[2].__contains__("М") and not data[2].__contains__("Д") \
                    and not data[1].__contains__("Капит"):
                data[3], data[4] = clear(data[3], "\xa0"), clear(data[4], "\xa0")
                big_data[-1]["Капитал. ремонт"] = f'Федеральный бюджет, тыс.руб.: {data[3]},' \
                                                  f" Областной бюджет, тыс.руб.: {data[4]}"
            elif data.__contains__("Оснащение"):
                data[3], data[4] = clear(data[3], "\xa0"), clear(data[4], "\xa0")
                big_data[-2]["Оснащение"] = f"Федеральный бюджет, тыс.руб.: {data[3]}," \
                                            f" Областной бюджет, тыс.руб.: {data[4]}"

    for i in big_data:
        if not i.__contains__("Адрес"):
            i["Адрес"] = clear(hand_data_txt.readline(), "\n")

    token = "beb2fd7eb442bc3a5c9078cfadada4a5c1dbcc0c"
    data = dt.Dadata(token)
    for i in range(len(big_data)-1):
        result = data.suggest("address", big_data[i]["Адрес"])
        if len(result) != 0:
            big_data[i]["Долгота"], big_data[i]["Широта"] = [result[0]["data"]['geo_lat'], result[0]["data"]['geo_lon']]

    big_data = pd.DataFrame(big_data)
    big_data.to_excel("table.xlsx")
