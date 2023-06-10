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
    4: 'Epic',     # (purple)
    5: 'Legendary' # (orange)
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

    if not row_data.get('patch') or int(row_data['patch'].split('.')[0]) > 3: return False
    if not row_data.get('inventoryType') or row_data['inventoryType'] not in ALLOWED_ITEM_INVENTORY_TYPES: return False
    if not row_data.get('displayInfoId'): return False
    if not row_data.get('quality') or row_data['quality'] < 1: return False
    if not row_data.get('itemSubClass'): return False

    ITEM_DATA = {
        'itemId': row_data['ItemID'],
        'displayId': row_data['displayInfoId'],
        'quality': row_data['quality'],
        'type': row_data['itemSubClass']
    }

    data[0][ALLOWED_ITEM_INVENTORY_TYPES[row_data['inventoryType']]].append(ITEM_DATA)
    return True


def get_column_names():
    # строка, колонка
    return [sheet[1][i].value for i in range(max_column)]


def create_data_list():
    rows_recorded = 0

    for row in sheet.iter_rows(min_row=2):
        row_data = {}
        for cell in row:
            if cell.value is None: continue
            row_data.update({column_names[cell.column-1]: cell.value})

        rows_recorded += append_items_to_data(row_data)

        if not rows_recorded % 1000:
            print('max-row:', max_row, 'rows-recorded:', rows_recorded)

    percent_diff = (rows_recorded * 100) / max_row
    print(f'\nИтемов было добавлено: {rows_recorded}')
    print(f'\nПроцент добавленных итемов от их общего количества: {percent_diff:.0f}%')


def write_data_to_json(file_name, data):
    with open(file_name, 'w', encoding='utf-8') as file:
        dump(data, file, ensure_ascii=False)


if __name__ == '__main__':

    wb = load_workbook('items.xlsx', read_only=True)
    sheet = wb[wb.sheetnames[0]]
    max_row = sheet.max_row
    max_column = sheet.max_column

    column_names = get_column_names()

    create_data_list()

    sorted_data = [{slot_id: sorted(item_data, reverse=True, key=lambda x: x['quality'])}
                   for slot_data in data for slot_id, item_data in slot_data.items()]

    write_data_to_json('itemsdata.json', sorted_data)

    print('\nПрограмма успешно завершилась!')
