import requests
from bs4 import BeautifulSoup
import json
import openpyxl

'''
<script type="text/javascript">
            var _ = g_items; _[47512]={"quality":4,"icon":"inv_jewelry_ring_44","name_ruru":"Исповедь грешника"};
            var _ = g_classes; _[5]={"name_ruru":"Жрец"};
            var _ = g_spells; _[47540]={"icon":"spell_holy_penance","name_ruru":"Исповедь"};
                              _[53005]={"icon":"spell_holy_penance","name_ruru":"Исповедь"};
                              _[53006]={"icon":"spell_holy_penance","name_ruru":"Исповедь"};
                              _[53007]={"icon":"spell_holy_penance","name_ruru":"Исповедь"};
                              _[54518]={"icon":"spell_holy_penance","name_ruru":"Исповедь"};
                              _[47666]={"icon":"spell_holy_penance","name_ruru":"Исповедь"};
                              _[47750]={"icon":"spell_holy_penance","name_ruru":"Исповедь"};
                              _[52983]={"icon":"spell_holy_penance","name_ruru":"Исповедь"};
                              _[52984]={"icon":"spell_holy_penance","name_ruru":"Исповедь"};
                              _[52985]={"icon":"spell_holy_penance","name_ruru":"Исповедь"};
                              _[52998]={"icon":"spell_holy_penance","name_ruru":"Исповедь"};
                              _[52999]={"icon":"spell_holy_penance","name_ruru":"Исповедь"};
                              _[53000]={"icon":"spell_holy_penance","name_ruru":"Исповедь"};
                              _[54520]={"icon":"spell_holy_penance","name_ruru":"Исповедь"};
                              _[66097]={"icon":"spell_nature_starfall","name_ruru":"Исповедь"};
                              _[66098]={"icon":"spell_holy_penance","name_ruru":"Исповедь"};
                              _[68029]={"icon":"spell_holy_penance","name_ruru":"Исповедь"};
                              _[68030]={"icon":"spell_holy_penance","name_ruru":"Исповедь"};
                              _[68031]={"icon":"spell_holy_penance","name_ruru":"Исповедь"};
                              _[69905]={"icon":"spell_holy_penance","name_ruru":"Исповедь"};
                              _[69906]={"icon":"spell_holy_penance","name_ruru":"Исповедь"};
                              _[71137]={"icon":"spell_holy_penance","name_ruru":"Исповедь"};
                              _[71138]={"icon":"spell_holy_penance","name_ruru":"Исповедь"};
                              _[71139]={"icon":"spell_holy_penance","name_ruru":"Исповедь"};
        </script>
'''

def make_xlsx_file(spell_dict, key):
    # print(spell_dict)

    wb = openpyxl.Workbook()
    wb.remove(wb['Sheet']) # удаляем дефолтную первую страницу в таблице

    wb.create_sheet(key)
    sheet = wb[key]

    rows = 0

    for spell in spell_dict:

        rows += 1

        for i in range(1, 2+1):

            # row сдвиг по вертикали column сдвиг по горизонтали
            cell = sheet.cell(row=rows, column=i)

            if i == 1:
                cell.value = spell['name_ruru']
            if i == 2:
                cell.value = spell['spell_id']

    wb.save(f'{key}.xlsx')

x_list = []

def make_dict_with_id(script_text, key):
    item_key = 'var _ = g_items'
    class_key = 'var _ = g_classes'
    spell_key = 'var _ = g_spells'

    if key == 'items':
        text = script_text[script_text.find(item_key)+len(item_key) : script_text.find(class_key)]
    elif key == 'spells':
        text = script_text[script_text.find(spell_key)+len(spell_key) : len(script_text)]

    text = text.split(';')

    text.remove(text[0])
    text.remove(text[-1])

    text_0 = text[0]

    body = text_0[text_0.find(']={')+2 : len(text_0)]

    spell_dict = json.loads(body)

    spell_id = text_0[text_0.find('_[')+2 : text_0.find(']={')]
    spell_dict.update({'spell_id':spell_id})

    print(f"{key} [{spell_dict['name_ruru']}] под id [{spell_dict['spell_id']}] успешно сохранён в файл.")

    x_list.append(spell_dict)

def get_html(objectName, key):
    # https://base.opiums.eu/?search=Исповедь#abilities

    url = 'https://base.opiums.eu/?search={name}#abilities'.format(name=objectName)

    source = requests.get(url)
    source_text = source.text

    soup = BeautifulSoup(source_text, 'html.parser')

    script = soup.find('body').find('script', type="text/javascript")
    # print(script)
    script_text = script.text
    # print(script_text)

    if len(script_text) == 1:
        print(f'!WARNING! {key} [{objectName}] не удалось найти в базе данных. !WARNING!')
        return

    make_dict_with_id(script_text, key)

def start_program():

    cat = input('Введите категорию для поиска: ')

    if cat != 'spells' and cat != 'items':
        print('\nВведите корректную категорию, spells либо items\n')
        start_program()
        return

    print('\n>>>>>>>>>>>>>>>>>>>> Операция началась <<<<<<<<<<<<<<<<<<<<')

    # object_list = ['Исповедь', 'Взрыв разума', 'Ментальный крик', 'Подавление боли', 'Заклинание для теста без имени', 'Придание сил', 'Левитация', 'Исчадие Тьмы']
    # object_list = ['Гримуар разгневанного гладиатора', 'Перстень постепенной регенерации', 'Плащ горящего заката', 'Вещь для теста без имени', 'Амулет костяного часового', 'Подвеска истинной крови']

    file = open(f'{cat}.txt', 'r', encoding='utf-8')
    object_list = file.readlines()

    for objectName in object_list:
        get_html(objectName, cat)

    file.close()

    make_xlsx_file(x_list, cat)

    print('>>>>>>>>>>>>>>>>>>>> Операция завершена <<<<<<<<<<<<<<<<<<<<\n')
    input('Нажмите ENTER для закрытия программы.\n')

if __name__ == '__main__':
    print('by Friskes.\n')
    print('Приветствую, данная программа умеет собирать id у предметов или\nзаклинаний по их названию и сохранять\
 полученные данные в xlsx файл.\n\nДля работы с программой запишите интересующие вас предметы в текстовый файл\n\
с названием items.txt либо заклинания в файл spells.txt затем перезайдите в\nпрограмму и введите ниже категорию\
 для поиска items или spells соответственно.\n')
    start_program()
