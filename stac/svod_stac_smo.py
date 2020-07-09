import modules

app = []
apres = [0, 0, 0, 0]

def replace_prof(result):
    try:
        file_format = modules.xml.dom.minidom.parse ('settings//profiles_stac.xml')
    except FileotFoundError:
        print ('Файл profiles_stac.xml не существует')
        return None
    zap_in_xml_Obj = file_format.getElementsByTagName('zap')
    for zap_in_xml in zap_in_xml_Obj:
        profil_in_xml_Obj = zap_in_xml.getElementsByTagName('prof')
        profil_in_xml = int(profil_in_xml_Obj[0].childNodes[0].data)
        for lpu in range(len(result)):
            for profil in range(1, len(result[lpu])):
                if result[lpu][profil][0] == profil_in_xml:
                    name_in_xml_Obj = zap_in_xml.getElementsByTagName('name')
                    name_in_xml = name_in_xml_Obj[0].childNodes[0].data
                    result[lpu][profil][0] = name_in_xml

def findeadd(result, lpu, prof, kod_usl, pac, kd_z, kd, summ):
    cell_lpu = cell_podr = cell_ksg = -1
    for ilpu in range (len(result)):
        if result[ilpu][0] == lpu:
            cell_lpu = ilpu
            break
    if cell_lpu == -1:
        result.append(list(app))
        cell_lpu = len(result) - 1
        result[cell_lpu].append(lpu)
        
    for iprof in range (1, len(result[cell_lpu])):
        if result[cell_lpu][iprof][0] == prof:
            cell_podr = iprof
            break
    if cell_podr == -1:
        result[cell_lpu].append(list(app))
        cell_podr = len(result[cell_lpu]) - 1
        result[cell_lpu][cell_podr].append(prof)
    for iksg in range (1, len(result[cell_lpu][cell_podr])):
        if result[cell_lpu][cell_podr][iksg][0] == kod_usl:
            cell_ksg = iksg
            break
    if cell_ksg == -1:
        result[cell_lpu][cell_podr].append(list(app))
        cell_ksg = len(result[cell_lpu][cell_podr]) - 1
        result[cell_lpu][cell_podr][cell_ksg].append(kod_usl)
        result[cell_lpu][cell_podr][cell_ksg].append(list(apres))
    result[cell_lpu][cell_podr][cell_ksg][1][0] += pac
    result[cell_lpu][cell_podr][cell_ksg][1][1] += kd_z
    result[cell_lpu][cell_podr][cell_ksg][1][2] += kd
    result[cell_lpu][cell_podr][cell_ksg][1][3] += summ

def calc_svod_smo(file_reports, n_svod, pre_result):
    '''Формирование сводного счета страховой компании по стационару, 
    входной параметр путь к файлу из ТФОМС, сводный счет'''
    
    sluch_Obj = file_reports.getElementsByTagName('SLUCH')
    month=int(file_reports.getElementsByTagName('MONTH')[0].childNodes[0].data)
    year = int(file_reports.getElementsByTagName('YEAR')[0].childNodes[0].data)
    id_smo = file_reports.getElementsByTagName("PLAT")[0].childNodes[0].data
    
    for sluch in sluch_Obj:
        svod = int(sluch.getElementsByTagName("NSVOD")[0].childNodes[0].data)
        if svod//100 == n_svod:        
            usl_Obj = sluch.getElementsByTagName('USL')
            for usl in usl_Obj:
                parse_res = [0]*7
                kd = None
                try:
                    kd = int(usl.getElementsByTagName("KD")[0].childNodes[0].data)
                except IndexError:
                    pass
                summ = modules.Decimal(usl.getElementsByTagName("SUMV_USL")[0].childNodes[0].data)
                if kd != None or summ != modules.Decimal(0):
                    parse_res[5] = kd
                    parse_res[4] = kd
                    kod_lpu_Obj = usl.getElementsByTagName('KODLPU')
                    parse_res[0] = int(kod_lpu_Obj[0].childNodes[0].data)
                    profil_Obj = usl.getElementsByTagName('PODR')
                    parse_res[1] = int(profil_Obj[0].childNodes[0].data)//10
                    kod_usl_Obj = usl.getElementsByTagName('CODE_USL')
                    parse_res[2] = kod_usl_Obj[0].childNodes[0].data

#                    summ = modules.Decimal(usl.getElementsByTagName("SUMV_USL")[0].childNodes[0].data)
                    parse_res[6] = summ
                    parse_res[3] = 1 if summ != 0 else 0
                    findeadd(pre_result, *parse_res)
                
    return id_smo, month, year

def svod_format(svod_out_sheet, format):
    idMO, MO = modules.parse.settings_MO()
    modules.outxlsx.set_format(svod_out_sheet, format)
    modules.outxlsx.set_page(svod_out_sheet, format)
    modules.outxlsx.set_width(svod_out_sheet, format)
    modules.outxlsx.set_height(svod_out_sheet, format)
    svod_out_sheet[format.getElementsByTagName('MO')[0].childNodes[0].data].value = MO
    svod_out_sheet[format.getElementsByTagName('MO_id')[0].childNodes[0].data].value = idMO

def svod_append(svod_out_sheet, format, row, result):
    column = 1
    ar_prof = [0, 0]
    ar_lpu = []
    for lpu in range(len(result)):
        ar_podr = []
        for prof in range(1, len(result[lpu])):
            ar_prof[0] = row
            for ksg in range (1, len(result[lpu][prof])):
                iter = 0
                rcell_Obj = format.getElementsByTagName('rcell')
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
                    if iter == 0:
                        svod_out_sheet [first_cell].value = column
                    elif iter == 1:    
                        svod_out_sheet [first_cell].value = result[lpu][0]
                    elif iter == 2:
                        svod_out_sheet [first_cell].value = result[lpu][prof][0]
                    elif iter == 3:
                        svod_out_sheet [first_cell].value = result[lpu][prof][ksg][1][0]
                    elif iter == 4:
                        svod_out_sheet [first_cell].value = result[lpu][prof][ksg][1][1]
                    elif iter == 5:
                        svod_out_sheet [first_cell].value = result[lpu][prof][ksg][1][2]
                    elif iter == 6:
                        svod_out_sheet [first_cell].value = result[lpu][prof][ksg][0]
                    elif iter == 7:
                        svod_out_sheet [first_cell].value = result[lpu][prof][ksg][1][3]
                    else:
                        print('Промах')
                    iter +=1
                column += 1
                ar_prof[1] = row
                row += 1

            #Предварительный итог по отделению
            iter = 0
            rcell_Obj = format.getElementsByTagName('pecell')
            for rcell in rcell_Obj:
                first_cell_Obj = modules.parse.parse(rcell, 'n_cell')
                try:
                    first_cell_letter = first_cell_Obj[0].childNodes[0].data
                    first_cell = first_cell_letter + str(row)
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
                if iter == 0:
                    value_Obj = modules.parse.parse(rcell, 'value')
                    svod_out_sheet [first_cell].value = value_Obj[0].childNodes[0].data + str(result[lpu][prof][0]) + ')'
                elif iter == 1:    
                    cell_summ = first_cell_letter + str(ar_prof[0]) + ':' + first_cell_letter + str(ar_prof[1])
                    svod_out_sheet [first_cell].value = '=sum(' + cell_summ + ')'
                    ar_podr.append(row)
                    #Число пациентов в отделении
                    pass
                elif iter == 2:
                    cell_summ = first_cell_letter + str(ar_prof[0]) + ':' + first_cell_letter + str(ar_prof[1])
                    svod_out_sheet [first_cell].value = '=sum(' + cell_summ + ')'
                    #Койко-дни фактические по отделению
                    pass
                elif iter == 3:
                    cell_summ = first_cell_letter + str(ar_prof[0]) + ':' + first_cell_letter + str(ar_prof[1])
                    svod_out_sheet [first_cell].value = '=sum(' + cell_summ + ')'
                    #Койко-дни учтенные по отделению
                    pass
                elif iter == 4:
                    pass
                elif iter == 5:
                    cell_summ = first_cell_letter + str(ar_prof[0]) + ':' + first_cell_letter + str(ar_prof[1])
                    svod_out_sheet [first_cell].value = '=sum(' + cell_summ + ')'
                    #Сумма по отделению
                    pass
                else:
                    print('Промах')
                iter +=1
            row += 1
        #Предварительный итог по подразделению
        iter = 0
        rcell_Obj = format.getElementsByTagName('pecell')
        for rcell in rcell_Obj:
            first_cell_Obj = modules.parse.parse(rcell, 'n_cell')
            try:
                first_cell_letter = first_cell_Obj[0].childNodes[0].data
                first_cell = first_cell_letter + str(row)
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
            if iter == 0:
                value_Obj = modules.parse.parse(rcell, 'value')
                svod_out_sheet [first_cell].value = value_Obj[0].childNodes[0].data + str(result[lpu][0]) + ')'
            elif iter == 1:
                cell_summ = '=' + first_cell_letter + str(ar_podr[0])
                for res in range(1, len(ar_podr)):
                    cell_summ += '+' + first_cell_letter + str(ar_podr[res])

                svod_out_sheet [first_cell].value = cell_summ
                ar_lpu.append(row)
                #Число пациентов в подразделению
                pass
            elif iter == 2:
                cell_summ = '=' + first_cell_letter + str(ar_podr[0])
                for res in range(1, len(ar_podr)):
                    cell_summ += '+' + first_cell_letter + str(ar_podr[res])

                svod_out_sheet [first_cell].value = cell_summ
                #Койко-дни фактические по подразделению
                pass
            elif iter == 3:
                cell_summ = '=' + first_cell_letter + str(ar_podr[0])
                for res in range(1, len(ar_podr)):
                    cell_summ += '+' + first_cell_letter + str(ar_podr[res])

                svod_out_sheet [first_cell].value = cell_summ
                #Койко-дни учтенные по подразделению
                pass
            elif iter == 4:
                pass
            elif iter == 5:
                cell_summ = '=' + first_cell_letter + str(ar_podr[0])
                for res in range(1, len(ar_podr)):
                    cell_summ += '+' + first_cell_letter + str(ar_podr[res])

                svod_out_sheet [first_cell].value = cell_summ
                #Сумма по подразделению
                pass
            else:
                print('Промах')
            iter +=1
        row += 1
    #Итог по больнице
    iter = 0
    rcell_Obj = format.getElementsByTagName('ecell')
    for rcell in rcell_Obj:
        first_cell_Obj = modules.parse.parse(rcell, 'n_cell')
        try:
            first_cell_letter = first_cell_Obj[0].childNodes[0].data
            first_cell = first_cell_letter + str(row)
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
        if iter == 0:
            value_Obj = modules.parse.parse(rcell, 'value')
            svod_out_sheet [first_cell].value = value_Obj[0].childNodes[0].data
        elif iter == 1:
            cell_summ = '=' + first_cell_letter + str(ar_lpu[0])
            for res in range(1, len(ar_lpu)):
                cell_summ += '+' + first_cell_letter + str(ar_lpu[res])

            svod_out_sheet [first_cell].value = cell_summ
            #Число пациентов в больнице
            pass
        elif iter == 2:
            cell_summ = '=' + first_cell_letter + str(ar_lpu[0])
            for res in range(1, len(ar_lpu)):
                cell_summ += '+' + first_cell_letter + str(ar_lpu[res])

            svod_out_sheet [first_cell].value = cell_summ
            #Койко-дни фактические по больнице
            pass
        elif iter == 3:
            cell_summ = '=' + first_cell_letter + str(ar_lpu[0])
            for res in range(1, len(ar_lpu)):
                cell_summ += '+' + first_cell_letter + str(ar_lpu[res])

            svod_out_sheet [first_cell].value = cell_summ
            #Койко-дни учтенные по больнице
            pass
        elif iter == 4:
            pass
        elif iter == 5:
            cell_summ = '=' + first_cell_letter + str(ar_lpu[0])
            for res in range(1, len(ar_lpu)):
                cell_summ += '+' + first_cell_letter + str(ar_lpu[res])

            svod_out_sheet [first_cell].value = cell_summ
            eres = '=' + first_cell
            #Сумма по больнице
            pass
        else:
            print('Промах')
        iter +=1
    row += 2
    #Завершение отчета
    iter = 0
    rcell_Obj = format.getElementsByTagName('edcell')
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
        value_Obj = modules.parse.parse(rcell, 'value')
        try:
            svod_out_sheet [first_cell].value = value_Obj[0].childNodes[0].data
        except IndexError:
            pass
        if iter == 1:
            svod_out_sheet [first_cell].value = eres
            row += 1
        iter += 1

def svod_smo_ks_ds(sourceDir, source, stac):
    #по страховым
    try:
        file_format = modules.xml.dom.minidom.parse ('settings//svod_stac_smo.xml')
    except FileotFoundError:
        print ('Файл svod_stac_smo.xml не существует')
        return None
    for i in range(len(source)):
        result = []
        svod_out_file = modules.openpyxl.Workbook ()
        svod_out_sheet = svod_out_file.active
        svod_out_sheet.title = 'Сводный'
        
        svod_format(svod_out_sheet, file_format)
        start_row = int(file_format.getElementsByTagName('start_row')[0].childNodes[0].data)
        
        smo, month, year = calc_svod_smo(source[i], stac, result)
        if result != []:
            result.sort()
            replace_prof(result)
            svod_append(svod_out_sheet, file_format, start_row, result)
            stac_stat = 'кс' if stac == 1 else 'дс'
            svod_out_file.save(sourceDir + modules.parse.delimiter + 'svod_stac_' + stac_stat + '_' + str(smo) +'.xlsx')
        svod_out_file.close

def svod_ks_ds_tfoms(sourceDir, source, stac):
#по ЛПУ с ТФОМС
    try:
        file_format = modules.xml.dom.minidom.parse ('settings//svod_stac_smo.xml')
    except FileotFoundError:
        print ('Файл svod_stac_smo.xml не существует')
        return None
    result = []
    for i in range(len(source)):
        svod_out_file = modules.openpyxl.Workbook ()
        svod_out_sheet = svod_out_file.active
        svod_out_sheet.title = 'Сводный'
        
        svod_format(svod_out_sheet, file_format)
        row = start_row = int(file_format.getElementsByTagName('start_row')[0].childNodes[0].data)
        
        smo, month, year = calc_svod_smo(source[i], stac, result)

    if result != []:
        result.sort()
        replace_prof(result)
        svod_append(svod_out_sheet, file_format, start_row, result)
        stac_stat = 'кс' if stac == 1 else 'дс'
        name_file = sourceDir + modules.parse.delimiter + 'svod_stac_' + stac_stat + '_с ТФОМС.xlsx'
        svod_out_file.save(name_file)
        svod_out_file.close

def svod_ks_ds(sourceDir, source, stac):
#по ЛПУ без ТФОМС
    try:
        file_format = modules.xml.dom.minidom.parse ('settings//svod_stac_smo.xml')
    except FileotFoundError:
        print ('Файл svod_stac_smo.xml не существует')
        return None
    result = []
    for i in range(len(source)):
        svod_out_file = modules.openpyxl.Workbook ()
        svod_out_sheet = svod_out_file.active
        svod_out_sheet.title = 'Сводный'
        
        svod_format(svod_out_sheet, file_format)
        row = start_row = int(file_format.getElementsByTagName('start_row')[0].childNodes[0].data)
        id_smo = source[i].getElementsByTagName("PLAT")[0].childNodes[0].data
        if id_smo != '61010':
            smo, month, year = calc_svod_smo(source[i], stac, result)

    if result != []:
        result.sort()
        replace_prof(result)
        svod_append(svod_out_sheet, file_format, start_row, result)
        stac_stat = 'кс' if stac == 1 else 'дс'
        svod_out_file.save(sourceDir + modules.parse.delimiter + 'svod_stac_' + stac_stat + '_без ТФОМС.xlsx')
    svod_out_file.close
