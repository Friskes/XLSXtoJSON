# pip install openpyxl
from openpyxl import load_workbook
from json import dump


ALLOWED_ITEM_INVENTORY_TYPES = {
    'Head': 1,
    'Shoulder': 3,
    'Shirt': 4,

    'Chest': 5,
    'Robe': 5,

    'Waist': 6,
    'Legs': 7,
    'Feet': 8,
    'Wrist': 9,
    'Hands': 10,
    'Cloak': 15,

    ##################################
    # 'Main Hand': 16, # Устаревший слот id?
    # 'One-Hand': 16,  # == 'Main Hand'
    # 'Two-Hand': 16,  # == 'Main Hand'

    # 'Off-Hand': 17,         # Устаревший слот id?
    # 'Held in Off-hand': 17, # == 'Off-Hand'
    # 'Shield': 17,           # == 'Off-Hand'
    ##################################

    'Ranged': 18,
    'Ranged Right': 18,
    'Thrown': 18,

    'Tabard': 19,

    ##################################
    'Main Hand': 21, # Новый слот id?
    'One-Hand': 21,  # == 'Main Hand'
    'Two-Hand': 21,  # == 'Main Hand'

    'Off-Hand': 22,         # Новый слот id?
    'Held in Off-hand': 22, # == 'Off-Hand'
    'Shield': 22            # == 'Off-Hand'
    ##################################
}

ITEM_QUALITIES = {
    0: 'Poor',     # (gray)
    1: 'Common',   # (white)
    2: 'Uncommon', # (green)
    3: 'Rare',     # (blue)
    4: 'Epic'      # (purple)
}

data = [{
    1: [],
    3: [],
    4: [],
    5: [],
    6: [],
    7: [],
    8: [],
    9: [],
    10: [],
    15: [],

    #######
    # 16: [],
    # 17: [],
    #######

    18: [],
    19: [],

    #######
    21: [],
    22: []
    #######
}]


def append_items_to_data(row_data):
    # {'ItemID': 17, 'ItemName': 'Martin Fury', 'quality': 0,
    # 'patch': '1.12.1', 'inventoryType': 'Shirt', 'displayInfoId': 5661}

    if not row_data.get('patch') or int(row_data['patch'].split('.')[0]) > 3: return
    if not row_data.get('inventoryType') or row_data['inventoryType'] not in ALLOWED_ITEM_INVENTORY_TYPES: return
    if not row_data.get('displayInfoId'): return
    if not row_data.get('quality') or row_data['quality'] < 2: return

    ITEM_DATA = {
        'itemId': row_data['ItemID'],
        'displayId': row_data['displayInfoId']
    }
    data[0][ALLOWED_ITEM_INVENTORY_TYPES[row_data['inventoryType']]].append(ITEM_DATA)

    # print(json.dumps(ITEM_DATA, indent=4, ensure_ascii=False))


def write_data_to_json(file_name):
    with open(file_name, 'w', encoding='utf-8') as file:
        dump(data, file, ensure_ascii=False)


def get_column_names():
    column_names = []
    for i in range(max_column):
        # строка, колонка
        column_names.append(sheet[1][i].value)
    return column_names


def create_data_list():

    row_recorded = 0
    for row in sheet.rows:
        row_data = {}
        for cell in row:
            try:
                if cell.column == 1:    # ItemID
                    row_data.update({column_names[0]: cell.value})
                elif cell.column == 2:  # ItemName
                    row_data.update({column_names[1]: cell.value})
                elif cell.column == 8:  # icon
                    row_data.update({column_names[7]: cell.value})
                elif cell.column == 11: # quality
                    row_data.update({column_names[10]: cell.value})
                elif cell.column == 12: # patch
                    row_data.update({column_names[11]: cell.value})
                elif cell.column == 41: # inventoryType
                    row_data.update({column_names[40]: cell.value})
                elif cell.column == 65: # displayInfoId
                    row_data.update({column_names[64]: cell.value})
            except AttributeError:
                continue

        row_recorded += 1
        if row_recorded == 1: continue

        append_items_to_data(row_data)

        print('max-row:', max_row, 'row-recorded:', row_recorded)

    print(f'{row_recorded} Итемов было добавлено в список data')


if __name__ == '__main__':
    wb = load_workbook('items.xlsx', read_only=True)
    sheet = wb[wb.sheetnames[0]]
    max_row = sheet.max_row
    max_column = sheet.max_column

    column_names = get_column_names()

    create_data_list()

    write_data_to_json('itemsdata.json')

    print('\nПрограмма успешно завершилась!')
