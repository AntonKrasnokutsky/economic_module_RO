import modules

months = ['января',
         'февраля',
         'марта',
         'апреля',
         'мая',
         'июня',
         'июля',
         'августа',
         'сентября',
         'октября',
         'ноября',
         'декабря']

month_now = ['январь',
             'февраль',
             'март',
             'апрель',
             'май',
             'июнь',
             'июль',
             'август',
             'сентябрь',
             'октябрь',
             'ноябрь',
             'декабрь']

def set_format(bill_sheet, nodes):
    cells = modules.parse.parse(nodes, 'cell')
    for node in cells:
        first_cell_Obj = modules.parse.parse(node, 'n_cell')
        try:
            first_cell = first_cell_Obj[0].childNodes[0].data
        except IndexError:
            break
        borders = modules.parse.border(node)
        bill_sheet [first_cell].border = modules.stiletxt.borders(borders)
        books = modules.parse.parse(node, 'mcell')
        end_cell = None
        for book in books:
            end_cell = book.childNodes[0].data
            bill_sheet [end_cell].border = modules.stiletxt.borders(borders)
        
        if end_cell != None:
            #объединение ячеек
            merge = str (first_cell + ':' + end_cell)
            #объединяем
            bill_sheet.merge_cells (merge)

        #Размер шрифта, жирность, курсив
        bill_sheet [first_cell].font = modules.stiletxt.font (*modules.parse.font(node))

        #выравнивание текста в ячейке (по горизонтали, по вертикали, перенос по словам)
        bill_sheet [first_cell].alignment = modules.stiletxt.alig(*modules.parse.aligment(node))
        
        number_format = modules.parse.parse(node, 'number_format')
        try:
            #формат ячейка, отображение дробных чисел
            bill_sheet [first_cell].number_format = number_format[0].childNodes[0].data
        except IndexError:
            pass
            
        text = modules.parse.parse(node, 'value')
        try:
            bill_sheet [first_cell].value = text[0].childNodes[0].data
        except IndexError:
            pass

def set_width(bill_sheet, nodes):
    columnObj = modules.parse.parse(nodes, 'letter')
    try:
        column = columnObj[0].childNodes[0].data
    except IndexError:
        pass
    else:
        columns = modules.parse.parse(nodes, 'column')
        for column in columns:
            column_letter_Obj = modules.parse.parse(column, 'letter')
            column_letter = column_letter_Obj[0].childNodes[0].data
            column_width_Obj = modules.parse.parse(column, 'width')
            column_width = int(column_width_Obj[0].childNodes[0].data)
            column_hidden_Obj = modules.parse.parse(column, 'hidden')
            column_hidden = int(column_hidden_Obj[0].childNodes[0].data)

            bill_sheet.column_dimensions [column_letter].width = column_width / 256
            bill_sheet.column_dimensions [column_letter].hidden = column_hidden

def set_height(bill_sheet, nodes):
    rowObj = modules.parse.parse(nodes, 'num')
    try:
        row = rowObj[0].childNodes[0].data
    except IndexError:
        pass
    else:
        rows = modules.parse.parse(nodes, 'row')
        for row in rows:
            row_num_Obj = modules.parse.parse(row, 'num')
            row_num = int(row_num_Obj[0].childNodes[0].data)
            row_height_Obj = modules.parse.parse(row, 'height')
            row_height = int(row_height_Obj[0].childNodes[0].data)
            row_hidden_Obj = modules.parse.parse(row, 'hidden')
            row_hidden = int(row_hidden_Obj[0].childNodes[0].data)

            bill_sheet.row_dimensions[row_num].height = row_height / 20
            bill_sheet.row_dimensions[row_num].hidden = row_hidden

def set_page(bill_sheet, nodes):
    p_left_Obj = modules.parse.parse(nodes, 'p_left')
    bill_sheet.page_margins.left =  float(p_left_Obj[0].childNodes[0].data)/2.54
    p_right_Obj = modules.parse.parse(nodes, 'p_right')
    bill_sheet.page_margins.right = float(p_right_Obj[0].childNodes[0].data)/2.54
    p_top_Obj = modules.parse.parse(nodes, 'p_top')
    bill_sheet.page_margins.top = float(p_top_Obj[0].childNodes[0].data)/2.54
    p_bottom_Obj = modules.parse.parse(nodes, 'p_bottom')
    bill_sheet.page_margins.bottom = float(p_bottom_Obj[0].childNodes[0].data)/2.54
    p_scale_Obj = modules.parse.parse(nodes, 'p_scale')
    bill_sheet.page_setup.scale = int(p_scale_Obj[0].childNodes[0].data)
    bill_sheet.page_setup.paperSize = '9'

def set_months(bill_sheet, nodes, id_smo, result):
    current_month_Obj = modules.parse.parse(nodes, 'current_month')
    try:
        bill_sheet[current_month_Obj[0].childNodes[0].data].value = month_now[int(id_smo[3])-1]
    except IndexError:
        pass
    else:
        past_month_Obj = modules.parse.parse(nodes, 'past_month')
        past_month = past_month_Obj[0].childNodes[0].data
        bill_sheet[past_month].value = month_now[int(id_smo[3])-2] if int(id_smo[3])-1 > 1 else month_now[11]
        bill_num_Obj = modules.parse.parse(nodes, 'bill_num')
        bill_num = bill_num_Obj[0].childNodes[0].data
        bill_sheet[bill_num].value = id_smo[1]
        bill_date_Obj = modules.parse.parse(nodes, 'bill_date')
        bill_date = bill_date_Obj[0].childNodes[0].data
        bill_sheet[bill_date].value = schet_date(id_smo)
        bill_year_Obj = modules.parse.parse(nodes, 'year')
        for bill_year in bill_year_Obj:
            bill_sheet[bill_year.childNodes[0].data].value = id_smo[4] + ' г.'
        set_page(bill_sheet, nodes)
        capitalize_Obj = modules.parse.parse(nodes, 'capitalize')
        capitalize = capitalize_Obj[0].childNodes[0].data
        #Переводим рубли в текст
        if id_smo[0] != '61010':
            value = modules.pytils.numeral.rubles (int(result [27]))
            #Отделяем рубли от копеек
            kop = int ((result [27] * 100 - int (result [27] ) * 100))
            bill_sheet[capitalize].value = value.capitalize() + ' ' + str (kop) + ' ' + modules.NumToStr.koptxt (kop)
        else:
            value = modules.pytils.numeral.rubles (int(result [14]) )
            #Отделяем рубли от копеек
            kop = int ((result [14] * 100 - int (result [14] ) * 100))
            bill_sheet[capitalize].value = value.capitalize() + ' ' + str (kop) + ' ' + modules.NumToStr.koptxt (kop)

def schet_date(date):
    if date[0] != '61010':
        result=date[2][8]+date[2][9] + '.' + date[2][5]+date[2][6] + '.'
        for i in range (4):
            result += date[2][i]
    else:
        result=date[2][8]+date[2][9] + ' ' + months[int(date[2][5]+date[2][6])-1] + ' '
        for i in range (4):
            result += date[2][i]
    return result