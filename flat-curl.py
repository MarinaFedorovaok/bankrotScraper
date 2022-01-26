import requests, json, sys, xlsxwriter
import PySimpleGUI as sg
import cookies as c

sg.theme('DarkAmber')   # Add a touch of color
# All the stuff inside your window.
layout = [  [sg.Text('Введите данные Вашей квартиры')],
            [sg.Text('Введите стоимость квартиры'), sg.InputText()],
            [sg.Text('Введите площадь квартиры в формате: 31.5 м2' ), sg.InputText()],
            [sg.Button('Ok'), sg.Button('Cancel')] ]

# Create the Window
window = sg.Window('Window Title', layout)
# Event Loop to process "events" and get the "values" of the inputs
while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Cancel': # if user closes window or clicks cancel
        break
       
window.close() #end of window 1

price = values[0]
area = values[1]
priceMetre = int(price)/float(area)

def make_request_and_wirite_it_down(locationId, rooms_nums_id):
    c.cookies
    c.headers 
    params = (
        ('key', 'af0deccbgcgidddjgnvljitntccdduijhdinfgjgfjir'),
        ('params[201]', '1059'), # Тип сделки -- покупка, аренда -- 1060
        #('metroId[]', ['154', '2132', '155']),  # Станции меторо, нужно узнать какие коды какой станции соответствуют.
        ('categoryId', '24'),   # категория товаров 24 -- квариры, 14 -- автомобили
        ('locationId', locationId), # регион,107621 -- СПб + ЛО.
        ('params[549][]', rooms_nums_id), 
        ('priceMin', '1000000'),
        ('priceMax', '5000000'),
        #('context', 'H4sIAAAAAAAAA0u0MrSqLraysFJKK8rPDUhMT1WyLrYys1LKzMvJzANyagF-_ClVIgAAAA'), # /n
        #  некоторая загадочноя строке, которая была в оригинальном запросе, но без которой все работает
        ('page', '1'),
        #('lastStamp', '1642688340'), #Похоже, что это time stamp запроса, без него все работает
        ('display', 'list'),
        ('limit', '50'),
    )

    res = requests.get('https://m.avito.ru/api/11/items', headers=c.headers, params=params, cookies=c.cookies)
    try:
        res = res.json()
    except json.decoder.JSONDecodeError:
        except_error(res)
    #print(res)
    if not (('status' in res) and (res['status'] == 'ok')):
        sys.stderr.write("responce status is not \'ok\'")
        sys.exit(1)
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
    priceMetre = 0
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
            print(title,'\t',area,'\t', price)
            sg.Print(title,'\t',area,'\t', price) #Debug Window
            worksheet.write(row, 0, title)
            worksheet.write(row, 1, float(area.replace(',','.')))
            worksheet.write(row, 2, int(price))
            worksheet.write(row, 3, '=C' + str(row +1) + '/B' + str(row +1))
            priceMetre = int(price)/float(area.replace(',','.'))
            print(priceMetre)
            priceMetreSumm = priceMetreSumm+priceMetre
            #totalMeters = totalMeters + float(area.replace(',','.'))
            #print(totalMeters)
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
locations = {'СПб': '653240', 'СПб + Ло': '107621', 'Самара': '653040', 'Калилинград': '630090'}
rooms_num = {'студия': '5695', '1 комната': '5696', '2 комнаты':'5697', '3 комнаты':'5698', '4 комнаты':'5699', '5 комнат':'5700', '6 комнат':'5701'}

#define layout
layout = [[sg.Text('Регион:',size=(20, 1), font = 'Lucida',justification = 'left')],
        [sg.Combo(list(locations.keys()), default_value='СПб', key = 'location')],
        [sg.Listbox(list(rooms_num.keys()), key = 'rooms_nums',\
                    no_scrollbar = True, select_mode = sg.LISTBOX_SELECT_MODE_MULTIPLE, size=(30, 7))],
        [sg.Button('OK', font = ('Times New Roman',12)), sg.Button('CANCEL', font = ('Times New Roman', 12))]]
#Define Window
win = sg.Window('avito-scraper',layout)
#Read  values entered by user
event, value = win.read()
#close first window
win.close()

#display string in a popup         
locationId = locations[value['location']]
print(value['rooms_nums'])

rooms_nums_id = [] # выбираем из словаря индексы квартир по описанию
for s in value['rooms_nums']:
    rooms_nums_id.append(rooms_num[s])
######################################
middle_price = make_request_and_wirite_it_down(locationId, rooms_nums_id)
profit_percent = (1-priceMetre/middle_price)*100
result1 = 'Средняя рыночная стоимость м2 = ' + str(round(middle_price, 2))
result2 = 'Средняя стоимость  м2 покупаемой квартиры = ' + str(round(priceMetre, 2))
result3 = 'Доходность = ' + str(round(profit_percent, 2)) + '%'
sg.popup(result1 + '\n' + result2 + '\n' + result3)

