from numpy import place
import requests, json, sys, xlsxwriter
import PySimpleGUI as sg
import cookies as c
from geopy.distance import distance
def count_items(params):
    count_params = params.copy()
    count_params['countOnly'] = '1'
    count_res = requests.get('https://m.avito.ru/api/11/items', headers=c.headers, params=count_params, cookies=c.cookies)
    try:
        count_res = count_res.json()
    except json.decoder.JSONDecodeError:
        except_error(count_res)
        sys.exit(1)
    print(count_res)
    if not (('status' in count_res) and (count_res['status'] == 'ok')):
        sys.stderr.write("count_items: responce status is not \'ok\'")
        sys.exit(1)
    return count_res['result']['count']
    
def make_request_and_wirite_it_down(locationId, rooms_nums_id, max_pages):
    items_on_page_limit = 50
    params = {
        'key': 'af0deccbgcgidddjgnvljitntccdduijhdinfgjgfjir',
        'params[201]': '1059', # Тип сделки -- покупка, аренда -- 1060
        #('metroId[]', ['154', '2132', '155']),  # Станции меторо, нужно узнать какие коды какой станции соответствуют.
        'categoryId': '24',   # категория товаров 24 -- квариры, 14 -- автомобили
        'locationId': locationId, # регион,107621 -- СПб + ЛО.
        'params[549][]': rooms_nums_id, 
        'priceMin': '1000000',
        'priceMax': '5000000',
        #('context', 'H4sIAAAAAAAAA0u0MrSqLraysFJKK8rPDUhMT1WyLrYys1LKzMvJzANyagF-_ClVIgAAAA'), # /n
        #  некоторая загадочноя строке, которая была в оригинальном запросе, но без которой все работает
        #('lastStamp', '1642688340'), #Похоже, что это time stamp запроса, без него все работает
        'display': 'list',
        'limit': str(items_on_page_limit),
    }
    ### Запрашиваем количество страниц
    items_num = int(count_items(params))
    print(items_num)
    pages = items_num//items_on_page_limit+1
    for i in range (1, min(max_pages,pages)):
        params['page'] = str(i)
        res = requests.get('https://m.avito.ru/api/11/items', headers=c.headers, params=params, cookies=c.cookies)
        try:
            res = res.json()
        except json.decoder.JSONDecodeError:
            except_error(res)
        #print(res)
        if not (('status' in res) and (res['status'] == 'ok')):
            sys.stderr.write("responce status is not \'ok\'")
            sys.exit(1)
        print(res)
        items = res['result']['items']
        workbook = xlsxwriter.Workbook('out.xlsx')
        worksheet = workbook.add_worksheet()

        row = 0
        worksheet.write(row, 0, "описание")
        worksheet.write(row, 1, "площадь")
        worksheet.write(row, 2, "цена")
        worksheet.write(row, 3, "стоимость м2")
        row += 1

        averagePrice = 0
        priceMetreSumm = 0

        for item in items:
            if item['type'] == 'item':
                value = item['value']
                splited_title = value['title'].split(', ')
                #print(splited_title)
                title = splited_title[0]
                area = splited_title[1] + splited_title[2]
                area = area.split('м')[0]
                price = value['price'].split('₽')[0].replace(' ','')
                coords = (float(value['coords']['lat']), float(value['coords']['lng']))
                place = (float(place_coord1), float(place_coord2))
                dist = distance(place, coords).km
                if dist < int(radius):
                    print(title,'\t',area,'\t', price)
                    sg.Print(title,'\t',area,'\t', price) #Debug Window
                    worksheet.write(row, 0, title)
                    worksheet.write(row, 1, float(area.replace(',','.')))
                    worksheet.write(row, 2, int(price))
                    worksheet.write(row, 3, '=C' + str(row +1) + '/B' + str(row +1))
                    priceMetre = int(price)/float(area.replace(',','.'))
                    print(priceMetre)
                    priceMetreSumm = priceMetreSumm+priceMetre
                    # totalMeters = totalMeters + float(area.replace(',','.'))
                    # print(totalMeters)
                    #print(totalPrice)
                    row += 1
    worksheet.write_formula(row, 3, '=AVERAGE(D' + str(2) + ':D' + str(row) + ')')
    averagePrice = priceMetreSumm / row
    print('Средняя стоимость m2:')
    print (averagePrice)
    workbook.close()
    return averagePrice

###########################################
###                 GUI                 ###
###########################################
# .Балашиха 2. Тюмень Екатеринбург Новосибирск Казань Воронеж г. Ростов-на-Дону


locations = {'СПб': '653240', 'СПб + Ло': '107621', 'Самара': '653040',\
     'Калилинград': '630090'}
rooms_num = {'студия': '5695', '1 комната': '5696', '2 комнаты':'5697', \
    '3 комнаты':'5698', '4 комнаты':'5699', '5 комнат':'5700', '6 комнат':'5701'}

#define layout

layout = [  [sg.Text('Введите данные Вашей квартиры')],
            [sg.Text('Введите стоимость квартиры'), sg.InputText('1000000', key = 'price')],
            [sg.Text('Введите площадь квартиры в формате: 31.5 м2' ), sg.InputText('31.5', key = 'area')],
            [sg.Text('Введите координаты Вашей квартиры в формате: 60.021946, 30.258681'), sg.Input('60.021946, 30.258681', key = 'place_text')],
            [sg.Text('Введите радиус поиска в формате: 10'), sg.InputText('10', key = 'radius')],
            [sg.Text('Количество просматриваемых страниц:'), sg.InputText('5', key = 'max_pages')],
            [sg.Text('Регион:', size=(20, 1), font = 'Lucida',justification = 'left')],
            [sg.Combo(list(locations.keys()), default_value='СПб', key = 'location')],
            [sg.Listbox(list(rooms_num.keys()), key = 'rooms_nums',\
                    no_scrollbar = True, select_mode = sg.LISTBOX_SELECT_MODE_MULTIPLE, size=(30, 7))],
            [sg.Button('OK', font = ('Times New Roman',12)), sg.Button('CANCEL', font = ('Times New Roman', 12))]]
#Define Window
win = sg.Window('avito-scraper',layout)
#Read  values entered by user
event, value = win.read()
#close first window

price = value['price']
area = value['area']
place_text = value['place_text']
radius = value['radius']
max_pages = int(value['max_pages'])
place_middle = place_text.find(',')
place_coord1 = place_text[:place_middle]
place_coord2 = place_text[(place_middle+1):]

win.close()

locationId = locations[value['location']]
print(value['rooms_nums'])

rooms_nums_id = [] # выбираем из словаря индексы квартир по описанию
for s in value['rooms_nums']:
    rooms_nums_id.append(rooms_num[s])

######################################

middle_price = make_request_and_wirite_it_down(locationId, rooms_nums_id, max_pages)
priceMetre = int(price)/float(area)
profit_percent = (middle_price/priceMetre - 1)*100
result1 = 'Средняя рыночная стоимость м2 = ' + str(round(middle_price, 2))
result2 = 'Средняя стоимость  м2 покупаемой квартиры = ' + str(round(priceMetre, 2))
result3 = 'Доходность = ' + str(round(profit_percent, 2)) + '%'
sg.popup(result1 + '\n' + result2 + '\n' + result3)

