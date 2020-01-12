import modules

prof_stomat = 8300

app = []
apres = [0]*6     

def replace_prof(result):
    try:
        file_format = modules.xml.dom.minidom.parse ('profiles_amb.xml')
    except FileotFoundError:
        print ('Файл profiles_amb.xml не существует')
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
                    continue

def findeadd(result, lpu, podr, kod_usl, kol_pas, kol_obr, kol_usl, kol_uet, kol_in_schet, summ):
    '''lpu - код ЛПУ в xml дом KODLPU
       podr - профиль подразделения в xml дом PODR 
       kod_usl - код услуги в xml дом CODE_USL в таривообразющей услуге <USL>IDSERV = IDMASTER</USL>
       kol_pas - колчиество посещений, количеству случаев с одной услугой (в коде услуги 5-я цифра "1" кроме диспансеризации и профосмотров)
       kol_obr - колчиество обращений, количеству случаев с несколькими услугами, кроме диспансеризации(в коде услуги 5-я цифра "2")
       kol_usl - колчество услуг
       kol_uet - количество УЕТ, соотношение стоимости услуги к тарифу услуги окруленное до второго знака после запятой SUMV_USL/TARIF
       kol_in_schet - колчиство индивидуальных счетов, колчисевто уникальных IDCASE
       summ - сумма за оказанную услугу, соответствует полю SUMV в доме SLUCH
    '''
    cell_lpu = cell_podr = cell_kod_usl = -1
    for ilpu in range (len(result)):
        if result[ilpu][0] == lpu:
            cell_lpu = ilpu
            break
    if cell_lpu == -1:
        result.append(list(app))
        cell_lpu = len(result) - 1
        result[cell_lpu].append(lpu)
        
    for ipodr in range (1, len(result[cell_lpu])):
        if result[cell_lpu][ipodr][0] == podr:
            cell_podr = ipodr
            break
    if cell_podr == -1:
        result[cell_lpu].append(list(app))
        cell_podr = len(result[cell_lpu]) - 1
        result[cell_lpu][cell_podr].append(podr)

    for ikod_usl in range (1, len(result[cell_lpu][cell_podr])):
        if result[cell_lpu][cell_podr][ikod_usl][0] == kod_usl:
            cell_kod_usl = ikod_usl
            break
    if cell_kod_usl == -1:
        result[cell_lpu][cell_podr].append(list(app))
        cell_kod_usl = len(result[cell_lpu][cell_podr]) - 1
        result[cell_lpu][cell_podr][cell_kod_usl].append(kod_usl)
        result[cell_lpu][cell_podr][cell_kod_usl].append(list(apres))
    result[cell_lpu][cell_podr][cell_kod_usl][1][0] += kol_pas
    result[cell_lpu][cell_podr][cell_kod_usl][1][1] += kol_obr
    result[cell_lpu][cell_podr][cell_kod_usl][1][2] += kol_usl
    result[cell_lpu][cell_podr][cell_kod_usl][1][3] += kol_uet
    result[cell_lpu][cell_podr][cell_kod_usl][1][4] += kol_in_schet
    result[cell_lpu][cell_podr][cell_kod_usl][1][5] += summ

def calc_svod_smo(file_reports, pre_result):
    '''Формирование сводного счета страховой компании по стационару, 
    входной параметр путь к файлу из ТФОМС, сводный счет'''
    n_svod = 3

    sluch_Obj = file_reports.getElementsByTagName('SLUCH')
    month=int(file_reports.getElementsByTagName('MONTH')[0].childNodes[0].data)
    year = int(file_reports.getElementsByTagName('YEAR')[0].childNodes[0].data)
    id_smo = file_reports.getElementsByTagName("PLAT")[0].childNodes[0].data
    
    for sluch in sluch_Obj:
        svod = int(sluch.getElementsByTagName("NSVOD")[0].childNodes[0].data)
        if svod//100 == n_svod:
            parse_res = [0]*9
            kol_usl = 0

            usl_in_xml_Obj = sluch.getElementsByTagName('USL')

            summ = modules.Decimal(sluch.getElementsByTagName("SUMV")[0].childNodes[0].data)
            parse_res[7] = 1 #Количество индивидуальных счетов
            parse_res[8] = summ #Сумма случая
            kod_lpu_in_xml_Obj = sluch.getElementsByTagName('KODLPU')
            #ЛПУ оказания
            parse_res[0] = int(kod_lpu_in_xml_Obj[0].childNodes[0].data)
            profil_in_xml_Obj = sluch.getElementsByTagName('PODR')
            #Подразделение оказания
            parse_res[1] = int(profil_in_xml_Obj[0].childNodes[0].data)
            
            for usl_in_xml in usl_in_xml_Obj:
                kol_usl += 1
                #try:
                #    parse_res[5] += int(sluch.getElementsByTagName("USL_OK")[0].childNodes[0].data)
                #except IndexError:
                #    pass
                id_master_in_xml_Obj = usl_in_xml_Obj[0].getElementsByTagName('IDMASTER')
                id_master_in_xml = id_master_in_xml_Obj[0].childNodes[0].data

                id_serv_in_xml_Obj = usl_in_xml.getElementsByTagName('IDSERV')
                id_serv_in_xml = id_serv_in_xml_Obj[0].childNodes[0].data

                #Количество услуг в случае
                parse_res[5] += 1
                if parse_res[1] == prof_stomat:
                    kod_usl_in_xml_Obj = usl_in_xml.getElementsByTagName('CODE_USL')
                    kod_usl_in_xml = kod_usl_in_xml_Obj[0].childNodes[0].data
                    #Код услуги
                    parse_res[2] = kod_usl_in_xml
                    #расчет УЕТ
                    sum_uls_in_xml_Obj = usl_in_xml.getElementsByTagName('SUMV_USL')
                    sum_uls_in_xml = sum_uls_in_xml_Obj[0].childNodes[0].data
                    tarif_in_xml_Obj = usl_in_xml.getElementsByTagName('TARIF')
                    tarif_in_xml = tarif_in_xml_Obj[0].childNodes[0].data
                    uet = float(sum_uls_in_xml) / float(tarif_in_xml)
                    parse_res[6] += modules.Decimal(int(uet*100)/100)
                    vid_pos = 3
                elif id_serv_in_xml == id_master_in_xml:
                    kod_usl_in_xml_Obj = usl_in_xml.getElementsByTagName('CODE_USL')
                    kod_usl_in_xml = kod_usl_in_xml_Obj[0].childNodes[0].data
                    #Код услуги
                    parse_res[2] = kod_usl_in_xml
                    vid_pos = int(int(kod_usl_in_xml)//1000000 - int(parse_res[1])*10)
                    vid_pos = 1 if vid_pos != 2 else 2
            
            if vid_pos == 1:
                parse_res[3] = 1
            elif vid_pos == 2:
                parse_res[4] = 1

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
        ar_prof[0] = row
        ar_podr = []
        for prof in range(1, len(result[lpu])):
            for code_usl in range (1, len(result[lpu][prof])):
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
                        svod_out_sheet [first_cell].value = result[lpu][0]
                    elif iter == 1:    
                        svod_out_sheet [first_cell].value = result[lpu][prof][0]
                    elif iter == 2:
                        svod_out_sheet [first_cell].value = result[lpu][prof][code_usl][0]
                    elif iter == 3:
                        svod_out_sheet [first_cell].value = result[lpu][prof][code_usl][1][0]
                    elif iter == 4:
                        svod_out_sheet [first_cell].value = result[lpu][prof][code_usl][1][1]
                    elif iter == 5:
                        svod_out_sheet [first_cell].value = result[lpu][prof][code_usl][1][2]
                    elif iter == 6:
                        svod_out_sheet [first_cell].value = result[lpu][prof][code_usl][1][3]
                    elif iter == 7:
                        svod_out_sheet [first_cell].value = result[lpu][prof][code_usl][1][4]
                    elif iter == 8:
                        svod_out_sheet [first_cell].value = result[lpu][prof][code_usl][1][5]
                    else:
                        print('Промах')
                    iter +=1
                column += 1
                ar_prof[1] = row
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
                svod_out_sheet [first_cell].value = value_Obj[0].childNodes[0].data + str(result[lpu][0])
            elif iter == 1:
                cell_summ = first_cell_letter + str(ar_prof[0]) + ':' + first_cell_letter + str(ar_prof[1])
                svod_out_sheet [first_cell].value = '=sum(' + cell_summ + ')'
                ar_lpu.append(row)
                #Количество посещений
                pass
            elif iter >= 2 and iter <= 6:
                cell_summ = first_cell_letter + str(ar_prof[0]) + ':' + first_cell_letter + str(ar_prof[1])
                svod_out_sheet [first_cell].value = '=sum(' + cell_summ + ')'
                #Количетво обращений
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
        elif iter >= 1 and iter <= 6:
            cell_summ = '=' + first_cell_letter + str(ar_lpu[0])
            for res in range(1, len(ar_lpu)):
                cell_summ += '+' + first_cell_letter + str(ar_lpu[res])

            svod_out_sheet [first_cell].value = cell_summ
            eres = '=' + first_cell
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

        iter += 1

def svod_smo_amb(sourceDir, source):
    #по страховым
    try:
        file_format = modules.xml.dom.minidom.parse ('svod_amb_smo.xml')
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
        
        smo, month, year = calc_svod_smo(source[i], result)
        if result != []:
            result.sort()
            replace_prof(result)
            svod_append(svod_out_sheet, file_format, start_row, result)
            svod_out_file.save(sourceDir + modules.parse.delimiter + 'svod_amb_' + str(smo) +'.xlsx')
        svod_out_file.close

def svod_amb_tfoms(sourceDir, source):
#по ЛПУ с ТФОМС
    try:
        file_format = modules.xml.dom.minidom.parse ('svod_amb_smo.xml')
    except FileotFoundError:
        print ('Файл svod_amb_smo.xml не существует')
        return None
    result = []
    for i in range(len(source)):
        svod_out_file = modules.openpyxl.Workbook ()
        svod_out_sheet = svod_out_file.active
        svod_out_sheet.title = 'Сводный'
        
        svod_format(svod_out_sheet, file_format)
        row = start_row = int(file_format.getElementsByTagName('start_row')[0].childNodes[0].data)
        
        smo, month, year = calc_svod_smo(source[i], result)

    if result != []:
        result.sort()
        replace_prof(result)
        svod_append(svod_out_sheet, file_format, start_row, result)
        name_file = sourceDir + modules.parse.delimiter + 'svod_amb'+ '_с ТФОМС.xlsx'
        svod_out_file.save(name_file)
        svod_out_file.close

def svod_amb(sourceDir, source):
#по ЛПУ без ТФОМС
    try:
        file_format = modules.xml.dom.minidom.parse ('svod_amb_smo.xml')
    except FileotFoundError:
        print ('Файл svod_amb_smo.xml не существует')
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
            smo, month, year = calc_svod_smo(source[i], result)

    if result != []:
        result.sort()
        replace_prof(result)
        svod_append(svod_out_sheet, file_format, start_row, result)
        svod_out_file.save(sourceDir + modules.parse.delimiter + 'svod_amb_' + 'без ТФОМС.xlsx')
    svod_out_file.close