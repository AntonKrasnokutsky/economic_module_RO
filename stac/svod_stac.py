import modules

def svod_format(svod_out_sheet, format):
    idMO, MO = modules.parse.settings_MO()
    modules.outxlsx.set_format(svod_out_sheet, format)
    modules.outxlsx.set_page(svod_out_sheet, format)
    modules.outxlsx.set_width(svod_out_sheet, format)
    modules.outxlsx.set_height(svod_out_sheet, format)
    svod_out_sheet[format.getElementsByTagName('MO')[0].childNodes[0].data].value = MO
    svod_out_sheet[format.getElementsByTagName('MO_id')[0].childNodes[0].data].value = idMO

def svod_append(svod_out_sheet, format, start_row, row, id_smo, result):
    rcell_Obj = format.getElementsByTagName('rcell')
    iter = 0
    for rcell in rcell_Obj:
        first_cell_Obj = modules.parse.parse(rcell, 'n_cell')
        try:
            first_cell = first_cell_Obj[0].childNodes[0].data + str(row)
        except IndexError:
            break
        borders = modules.parse.border(rcell)
        svod_out_sheet[first_cell].border = modules.stiletxt.borders(borders)
        books = modules.parse.parse(rcell, 'mcell')
        end_cell = None
        for book in books:
            end_cell = book.childNodes[0].data + str(row)
            svod_out_sheet[end_cell].border = modules.stiletxt.borders(borders)
        if end_cell != None:
            #объединение ячеек
            merge = str (first_cell + ':' + end_cell)
            #объединяем
            svod_out_sheet.merge_cells (merge)
        svod_out_sheet [first_cell].font = modules.stiletxt.font (*modules.parse.font(rcell))
        #выравнивание текста в ячейке (по горизонтали, по вертикали, перенос по словам)
        svod_out_sheet [first_cell].alignment = modules.stiletxt.alig(*modules.parse.aligment(rcell))
        
        number_format = modules.parse.parse(rcell, 'number_format')
        try:
            #формат ячейка, отображение дробных чисел
            svod_out_sheet [first_cell].number_format = number_format[0].childNodes[0].data
        except IndexError:
            pass
        pass
        if iter == 0:
            svod_out_sheet [first_cell].value = row - start_row + 1
        elif iter == 1:    
            svod_out_sheet [first_cell].value = id_smo
        else:
            svod_out_sheet [first_cell].value = result[iter-2]
        iter +=1

def svod_close(svod_out_sheet, format, start_row, end_row, month, year):
    cell_month_Obj = format.getElementsByTagName('month')
    for cell_month in cell_month_Obj:
        svod_out_sheet[cell_month.childNodes[0].data].value = month
    cell_year_Obj = format.getElementsByTagName('year')
    for cell_year in cell_year_Obj:
        svod_out_sheet[cell_year.childNodes[0].data].value = str(year) + ' г.'
    cell_month_year_Obj = format.getElementsByTagName('month_year')
    for cell_month_year in cell_month_year_Obj:
        svod_out_sheet[cell_month_year.childNodes[0].data].value = modules.outxlsx.month_now[month-1] + ' ' + str(year) + ' г.'

    svod_out_sheet[format.getElementsByTagName('svod_num')[0].childNodes[0].data].value = month
    rcell_Obj = format.getElementsByTagName('rcell')
    iter = 0
    for rcell in rcell_Obj:
        first_cell_Obj = modules.parse.parse(rcell, 'n_cell')
        try:
            first_letter = first_cell_Obj[0].childNodes[0].data
            first_cell = first_letter + str(end_row)
        except IndexError:
            break
        borders = modules.parse.border(rcell)
        svod_out_sheet[first_cell].border = modules.stiletxt.borders(borders)
        books = modules.parse.parse(rcell, 'mcell')
        end_cell = None
        for book in books:
            end_cell = book.childNodes[0].data + str(end_row)
            svod_out_sheet[end_cell].border = modules.stiletxt.borders(borders)
        if end_cell != None:
            #объединение ячеек
            merge = str (first_cell + ':' + end_cell)
            #объединяем
            svod_out_sheet.merge_cells (merge)
        svod_out_sheet [first_cell].font = modules.stiletxt.font (*modules.parse.font(rcell))
        #выравнивание текста в ячейке (по горизонтали, по вертикали, перенос по словам)
        svod_out_sheet [first_cell].alignment = modules.stiletxt.alig(*modules.parse.aligment(rcell))
        
        number_format = modules.parse.parse(rcell, 'number_format')
        try:
            #формат ячейка, отображение дробных чисел
            svod_out_sheet [first_cell].number_format = number_format[0].childNodes[0].data
        except IndexError:
            pass
        pass
        if iter == 0:
            svod_out_sheet [first_cell].value = 'Всего'
        else:
            svod_out_sheet [first_cell].value = '=sum(' + first_letter + str(start_row) + ':' + first_letter + str(end_row-1) + ')'
        iter +=1

def calc_svod(file_reports, n_svod):
    '''Формирование сводного счета, входной параметр путь к файлу из ТФОМС'''

    sluch_Obj = file_reports.getElementsByTagName('SLUCH')
    month=int(file_reports.getElementsByTagName('MONTH')[0].childNodes[0].data)
    year = int(file_reports.getElementsByTagName('YEAR')[0].childNodes[0].data)
    id_smo = file_reports.getElementsByTagName("PLAT")[0].childNodes[0].data
    
    result = [0]*4
    pacient = 0
    for sluch in sluch_Obj:
        svod = int(sluch.getElementsByTagName("NSVOD")[0].childNodes[0].data)
        if svod//100 == n_svod:        #Амболаторная помощь
            
            #Фактическое количесво дней в стационаре
            result[1] += int(sluch.getElementsByTagName("KD_Z")[0].childNodes[0].data)
            
            usl_Obj = sluch.getElementsByTagName('USL')
            kd = 0
            for usl in usl_Obj:
                try:
                    kd += int(usl.getElementsByTagName("KD")[0].childNodes[0].data)
                except IndexError:
                    pass

                summ = modules.Decimal(usl.getElementsByTagName("SUMV_USL")[0].childNodes[0].data)
                if summ != 0:
                    result[3] += summ
                    pacient += 1    
            result[2] += kd

    result[0] = pacient
    return id_smo, result, month, year

def svod_ks_ds_tfoms(sourceDir, source, stac):
    #с ТФОМС
    svod_out_file = modules.openpyxl.Workbook ()
    svod_out_sheet = svod_out_file.active
    svod_out_sheet.title = 'Сводный'
    try:
        file_format = modules.xml.dom.minidom.parse ('settings//svod_stac.xml')
    except FileNotFoundError:
        print ('Файл svod_stac.xml не существует')
        return None
   
    svod_format(svod_out_sheet, file_format)
    row = start_row = int(file_format.getElementsByTagName('start_row')[0].childNodes[0].data)
    
    for i in range(len(source)):
        smo, result, month, year = calc_svod(source[i], stac)
        if result[3] != 0:
            svod_append(svod_out_sheet, file_format, start_row, row, smo, result)
            row +=1

    svod_close(svod_out_sheet, file_format, start_row, row, month, year)
    stac_stat = 'кс' if stac == 1 else 'дс'
    svod_out_file.save(sourceDir + modules.parse.delimiter + 'svod_stac_smo_' + stac_stat +'_с ТФОМС.xlsx')
    svod_out_file.close

def svod_ks_ds(sourceDir, source, stac):
    #без ТФОМС
    svod_out_file = modules.openpyxl.Workbook ()
    svod_out_sheet = svod_out_file.active
    svod_out_sheet.title = 'Сводный'
    try:
        file_format = modules.xml.dom.minidom.parse ('settings//svod_stac.xml')
    except FileNotFoundError:
        print ('Файл svod_stac.xml не существует')
        return None
   
    svod_format(svod_out_sheet, file_format)
    row = start_row = int(file_format.getElementsByTagName('start_row')[0].childNodes[0].data)
    
    for i in range(len(source)):
        id_smo = source[i].getElementsByTagName("PLAT")[0].childNodes[0].data
        if id_smo != '61010':
            smo, result, month, year = calc_svod(source[i], stac)
            if result[3] != 0:
                svod_append(svod_out_sheet, file_format, start_row, row, smo, result)
                row +=1

    svod_close(svod_out_sheet, file_format, start_row, row, month, year)
    stac_stat = 'кс' if stac == 1 else 'дс'
    svod_out_file.save(sourceDir + modules.parse.delimiter + 'svod_stac_smo_' + stac_stat +'_без ТФОМС.xlsx')
    svod_out_file.close