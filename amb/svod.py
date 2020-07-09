import modules

profil = ['ks',
          'ds',
          'ambu',
          'smp']

def calc_svod(file_reports):
    '''Формирование сводного счета, входной параметр путь к файлу из ТФОМС'''
    sluch_Obj = file_reports.getElementsByTagName('SLUCH')
    month=int(file_reports.getElementsByTagName('MONTH')[0].childNodes[0].data)
    year = int(file_reports.getElementsByTagName('YEAR')[0].childNodes[0].data)
    id_smo = file_reports.getElementsByTagName("PLAT")[0].childNodes[0].data
    
#    disp_out_file = modules.openpyxl.Workbook ()
#    disp_out_sheet = disp_out_file.active
#    rows = 1
    result = [0]*5
    for sluch in sluch_Obj:
        uslInSluch = 0
        svod = int(sluch.getElementsByTagName("NSVOD")[0].childNodes[0].data)
        summ = modules.Decimal(sluch.getElementsByTagName("SUMV")[0].childNodes[0].data)
        if svod//100 == 3:        #Амболаторная помощь
            result[4] += summ
            usl_Obj = sluch.getElementsByTagName('USL')
            for usl in usl_Obj:
                uslInSluch += 1
            try:
                disp = sluch.getElementsByTagName("DISP")[0].childNodes[0].data
            except IndexError:
                disp = None
            profile = int(sluch.getElementsByTagName("PROFIL")[0].childNodes[0].data)
            if profile >= 85 and profile <= 90:
                usl_Obj = sluch.getElementsByTagName('USL')
                for usl in usl_Obj:
                    summ_usl = float(sluch.getElementsByTagName("SUMV_USL")[0].childNodes[0].data)
                    tarif = float(sluch.getElementsByTagName("TARIF")[0].childNodes[0].data)
                    uet = int(summ_usl/tarif*100)
                    result[3] += modules.Decimal(uet)/100
            elif uslInSluch == 1:
                result[0] += 1
            elif disp != None:
                #print(disp)
                result[0] += 1
            else:
                result[1] += 1 if uslInSluch != 0 else 0
                result[2] += uslInSluch
            idcase = sluch.getElementsByTagName("IDCASE")[0].childNodes[0].data
            #print(idcase, disp)
            cod_usl_Obj = sluch.getElementsByTagName("CODE_USL")
#            for cod_usl in cod_usl_Obj:
#                disp_out_sheet.cell(row = rows, column = 1).value = sluch.getElementsByTagName("IDCASE")[0].childNodes[0].data
#                disp_out_sheet.cell(row = rows, column = 2).value = disp
#                disp_out_sheet.cell(row = rows, column = 3).value = cod_usl.childNodes[0].data 
#                rows += 1

#    disp_out_file.save('d:\\1\\disp' + id_smo + '.xlsx')
    return id_smo, result, month, year

def svod_format(svod_out_sheet, format):
    idMO, MO = modules.parse.settings_MO()
    modules.outxlsx.set_format(svod_out_sheet, format)
    modules.outxlsx.set_width(svod_out_sheet, format)
    modules.outxlsx.set_height(svod_out_sheet, format)
    svod_out_sheet[format.getElementsByTagName('MO')[0].childNodes[0].data].value = MO
    svod_out_sheet[format.getElementsByTagName('MO_id')[0].childNodes[0].data].value = idMO

def svod_append(svod_out_sheet, format, row, id_smo, result):
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
            svod_out_sheet [first_cell].value = id_smo
        else:
            svod_out_sheet [first_cell].value = result[iter-1]
        iter +=1

def svod_close(svod_out_sheet, format, start_row, end_row, month, year):
    svod_out_sheet[format.getElementsByTagName('month')[0].childNodes[0].data].value = month
    svod_out_sheet[format.getElementsByTagName('year')[0].childNodes[0].data].value = str(year) + ' г.'
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
    

def svod_amb(sourceDir, source):
    svod_out_file = modules.openpyxl.Workbook ()
    svod_out_sheet = svod_out_file.active
    svod_out_sheet.title = 'Сводный'
    try:
        file_format = modules.xml.dom.minidom.parse ('settings//svod_amb.xml')
    except FileotFoundError:
        print ('Файл svod_amb.xml не существует')
        return None
    svod_format(svod_out_sheet, file_format)
    row = start_row = int(file_format.getElementsByTagName('start_row')[0].childNodes[0].data)

    for i in range(len(source)):
        smo, result, month, year = calc_svod(source[i])
        svod_append(svod_out_sheet, file_format, row, smo, result)
        row +=1
    svod_close(svod_out_sheet, file_format, start_row, row, month, year)
    svod_out_file.save(sourceDir + modules.parse.delimiter + 'svod_amb.xlsx')
    svod_out_file.close