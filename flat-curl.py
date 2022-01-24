import requests, json, sys, xlsxwriter
import PySimpleGUI as sg

""""
sg.theme('DarkAmber')   # Add a touch of color
# All the stuff inside your window.
layout = [  [sg.Text('Введите данные Вашей квартиры')],
            [sg.Text('Введите стоимость квартиры'), sg.InputText()],
            [sg.Text('Введите площадь квартиры'), sg.InputText()],
            [sg.Button('Ok'), sg.Button('Cancel')] ]

# Create the Window
window = sg.Window('Window Title', layout)
# Event Loop to process "events" and get the "values" of the inputs
while True:
    event, values = window.read()
    if event == sg.WIN_CLOSED or event == 'Cancel': # if user closes window or clicks cancel
        break
    print('You entered ', values[0], values [1])
window.close() #end of window 1

price = values[0]
area = values[1]
"""
priceMetre = 0#int(price)/float(area)

def make_request_and_wirite_it_down(locationId, rooms_nums_id):
    cookies = {
        'u': '2t4du0g6.qdwuhb.a5e8mudzyg80',
        'buyer_laas_location': '653240',
        'luri': 'sankt-peterburg',
        'buyer_location_id': '653240',
        'sx': 'H4sIAAAAAAAC%2F1TQXY7iMAwA4LvkmYf82jG3aWIb2sJsWzpp6Yi7r3YlpOEC38P3YwAAKiMoASWIQIJFAjEmWysymfOPaeZsuMO5xdVqud%2BOabiud35s2yIXiX7WYE5GzNlB9IgxWfs6GYSEUarrkAmkI3aWVERd1YqW%2FVuWcQ9X%2FZ7tgPWxU%2FH7sHFZ%2B01ae6zrLxky5P%2ByepWaIlDBSOiEIXEunD3GkhK%2B5dtlfNZpmcdld%2Fm42qHtt9kH22yen3H7kNHi62SqV%2BocCVnHLrPtEnsIIKyV1Kf8lvvndpmH4ehJjv7rumcKKd1xGR843ZflcyPAPzllL8UG7EAzcAgxusqszkXsMuhbPoro9AXzEi7NTZD%2BRO1kuuT63Q%2BtjR8yRni9%2FgYAAP%2F%2FMkczy8MBAAA%3D',
        '_ga_9E363E7BES': 'GS1.1.1642688110.5.1.1642688393.32',
        '_gcl_au': '1.1.1159494156.1642679970',
        '_ga': 'GA1.2.825310499.1642679972',
        '_ym_uid': '1642417546688206926',
        '_ym_d': '1642679976',
        '_gid': 'GA1.2.773285384.1642679977',
        '_ym_isad': '2',
        'f': '5.df155a60305e515a4b5abdd419952845a68643d4d8df96e9a68643d4d8df96e9a68643d4d8df96e9a68643d4d8df96e94f9572e6986d0c624f9572e6986d0c624f9572e6986d0c62ba029cd346349f36c1e8912fd5a48d02c1e8912fd5a48d0246b8ae4e81acb9fa1a2a574992f83a9246b8ae4e81acb9fad99271d186dc1cd0e992ad2cc54b8aa8b175a5db148b56e9bcc8809df8ce07f640e3fb81381f3591956cdff3d4067aa559b49948619279117b0d53c7afc06d0b2ebf3cb6fd35a0acba0ac8037e2b74f90df103df0c26013a7b0d53c7afc06d0bba0ac8037e2b74f9f722fe85c94f7d0c71e7cb57bbcb8e0f268a7bf63aa148d220f3d16ad0b1c5460df103df0c26013a03c77801b122405c868aff1d7654931c9d8e6ff57b051a588ad0bd3a9e57c5ef73122486354233fb938bf52c98d70e5c5f1a4d4589b2239d7b15a3f4ebc103c7d21ab7cd585086e0fecc7e6bc4080c2f1902ae0084209fc5e2415097439d404746b8ae4e81acb9fa786047a80c779d5146b8ae4e81acb9fa90bf83d8184497502da10fb74cac1eab2da10fb74cac1eabd1d953d27484fd81666d5156b5a01ea6',
        'ft': '19I1pLxrGYCqQbx6EItjuCFMUGwxTWyhMLTafhrim4MxOAx4LIIiwahDX4iPk8fRyDOW2ee46rm+MIUy6wdDyaM21MicX4MS1N8rB8NVA/1CRKZpdtdxvIAIHsLanGF0sP+rngSYClgI9MSoDwkiPAzxvJ2EFGH7i3Dbt3FsVXwlhPYKTE9MQTMo74QKgvOe',
        '__gads': 'ID=e4334bc798e493f5-2219475625cd0091:T=1642682324:RT=1642688373:S=ALNI_MYj-1Yagq9Mcv7NeLTx2PZzG2Q0Mw',
        'v': '1642688100',
        'dfp_group': '5',
        '_ym_visorc': 'b',
        '_mlocation': '621540',
        '_mlocation_mode': 'default',
        '_dc_gtm_UA-2546784-1': '1',
        'st': 'eyJhbGciOiJIUzI1NiIsInR5cCI6IkpXVCJ9.eyJkYXRhIjoidko0TkJ5SlU2TElJTGRDY3VNTkdBZm5PRWwySzdtS01KZXBFMDh1VUFIS3Z2dEZQcnVFOHlvZmFPc3RDdkRnSkFtM0RFTVZyRCtiaVlVb3dnVURGazgxbGZ1b2I0eDBlQ1F0ZUdKTENnc0FtNXFTTFNGSWsrWHdMZDM2VGpaMWozSGdMMlBUdGhZVVI0MTJHTVg1UGUxU2ZCZ0ZTYTZJS3o2Q2J2bDNGMWZrcDcrWnFwblZNY25jUko2UGJtbE56M3pwNXovYTRMTzJUelhVRzY0UzQ3WHM5NVJhemlOcnAvNmEycU5rMDh0dW05eXp1YjhwUFJXRythcGdEWk5yMDBkMDNzekhBa1c3THEwQThuUnUwS1cvYmZiejFxTUhTeFhPTVNpVlJ5Q2Mxb3V4dXNOUUs2VlZWd0tpTitZZ0toZkpSQjJ3ZFVIbWJaT000b2RrUXJhckRPV2VWZ285VFdSaGxqTjlzMUpScFZVWjdzMVZzd1h4aXBNazVWSFIvNHliZFM1UDRjUDl5SG1WL1N4RitndVdCbmNQemxKVW5aWEtrR2ZHY0N5OHZ4UFR6M0p1VjdGWmxPaUltWktIVUZab1hwZWNVQnp0eEUzdTlLaVZRN1FxdFNlVGJJbXNZQVdFdXB1MlRrUk5ZaERJUXJJNmN1ZVgzRnpLbVQ2MUR3YWVJVHh6ejNwbE9BUkZBcTJxWk9OazNSaU5MV3NiODhXSG9iNmlzSGJJRlZFMjJnWDA2UFJoSWw5MzBEeDdQMW5VV1dIQzA2aGtMdXl2Mm9neUdIRTN5S25YVlZMWUc1WDBZckQ5ZjdiajNGNGdtQzJBM2ZHME02UmhzZFNnSVhjTjlwLzNTMGVndzJQTTdmOTlEdnhrVUE3ZU9uekd0VDljTG1VYjl0WU9XUzRzVVdoSHVEamErNlUwaE1OMkZBQ3p4bXBqR1E2SUx4cTdLVkwzTk1GdXI3ODl2bjI3T2Q5aVpHSWFlTzlpMUwxNkF1QzFaTHJ0UWtBODVIZy85bFZGV2NhWWxWandNMzMrd2tBQnYyWkVnUkp3ZjNhNXJuNTJ6cTJMQURWQnIzZk9COUMyWG1IMzdyRVNrVjhOM0c0TEdnakNWalFoUDg4VVY5cXplOUI2TTY0TVVkREEzWEF5NmhmTE81NkRqdVZ1bWRCMkkzQjVuZEdBSnM5c3Q0QlF6dUJiR24zY08ySUF4b3Axc1BKd1pQQlNKbnM4NUNhSDF4dDBtRVZVPSIsImlhdCI6MTY0MjY4MDAwMSwiZXhwIjoxNjQzODg5NjAxfQ.DP2jwbqMYR1iYSh7zKPckzWBhMn-4yNwQy7d85wA-T4',
        'ST-TEST': 'TEST',
    }

    headers = {
        'User-Agent': 'Mozilla/5.0 (iPhone; CPU iPhone OS 14_6 like Mac OS X) AppleWebKit/605.1.15 (KHTML, like Gecko) Version/14.0.3 Mobile/15E148 Safari/604.1',
        'Accept': 'application/json, text/plain, */*',
        'Accept-Language': 'ru-RU,ru;q=0.8,en-US;q=0.5,en;q=0.3',
        'Accept-Encoding': 'gzip, deflate, br',
        'Content-Type': 'application/json;charset=utf-8',
        'Connection': 'keep-alive',
        'Referer': 'https://m.avito.ru/items/search',
        'Sec-Fetch-Dest': 'empty',
        'Sec-Fetch-Mode': 'cors',
        'Sec-Fetch-Site': 'same-origin',
        'TE': 'trailers',
    }

    params = (
        ('key', 'af0deccbgcgidddjgnvljitntccdduijhdinfgjgfjir'),
        ('params[201]', '1059'), # Тип сделки -- покупка, аренда -- 1060
        #('metroId[]', ['154', '2132', '155']),  # Станции меторо, нужно узнать какие коды какой станции соответствуют.
        ('categoryId', '24'),   # категория товаров 24 -- квариры, 14 -- автомобили
        ('locationId', locationId), # регион,107621 -- СПб + ЛО, нужно еще раз проверить.
        ('params[549][]', rooms_nums_id), 
        ('priceMin', '1000000'),
        ('priceMax', '5000000'),
        #('context', 'H4sIAAAAAAAAA0u0MrSqLraysFJKK8rPDUhMT1WyLrYys1LKzMvJzANyagF-_ClVIgAAAA'), # некоторая загадочноя строке, которая была в оригинальном запросе, но без которой все работает
        ('page', '1'),
        #('lastStamp', '1642688340'), #Похоже, что это time stamp запроса, без него все работает
        ('display', 'list'),
        ('limit', '50'),
    )

    res = requests.get('https://m.avito.ru/api/11/items', headers=headers, params=params, cookies=cookies)
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

locations = {'СПб': '653240', 'СПб + Ло': '107621'}
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
result1 = 'Средняя рыночная стоимость м2 = ' + str(round(make_request_and_wirite_it_down(locationId, rooms_nums_id),2))
result2 = 'Средняя стоимость  м2 покупаемой квартиры = ' + str(round(priceMetre))
sg.popup(result1 + '\n' + result2)

