import modules

delimiter = '/'
dispancer_adult_I = 'ДВ4'
dispancer_adult_II = 'ДВ2'
profosmotr_adult = 'ОПВ'
dispancer_minors_difficult_I = 'ДС1'
dispancer_minors_difficult_II = 'ДС3'
dispancer_minors_care_I = 'ДС2'
dispancer_minors_care_II = 'ДС3'
medical_exam_minors_I = 'ПН1'
medical_exam_minors_II = 'ПН2'

def value(SourcePath:str, tag:str, index=0):
    '''Ищем в файле тег и возращаем его значение
SouurcePath - путь к файлу, tag - тэг, значение которого необходимо вернуть'''
    try:
        SourceFile = modules.xml.dom.minidom.parse (SourcePath)
    except FileNotFoundError:
        return 'Файл '+ SourcePath + ' не существует'

    SourceFile.normalize()
    try:
        search_value = SourceFile.getElementsByTagName(tag)[index]
    except IndexError:
        return 'Файл не содержит тэг: ' + tag
    else:
        return search_value.childNodes[index].nodeValue

def data(SourceFile, tag:str):
    try:
        parseObj = SourceFile.getElementsByTagName(tag)[0]
    except IndexError:
        return False
    else:
        return parseObj.childNodes[0].data

def parse(SourceFile, tag:str):
    '''Ищем в файле все записи с искомым тегом, возвращаем список с индексами 
SouurcePath - путь к файлу, tag - искомый тег'''
    try:
        parseObj = SourceFile.getElementsByTagName(tag)
    except IndexError:
        return False
    else:
        return parseObj


def border(book):
    borders = []
    in_border = []
    in_border.append(parse(book, 'left'))
    in_border.append(parse(book, 'right'))
    in_border.append(parse(book, 'top'))
    in_border.append(parse(book, 'bottom'))
    for bor in in_border:
        try:
            border = bor[0].childNodes[0].data
        except IndexError:
            borders.append(None)
        else:
            borders.append(border)
    return borders

def font(book):
    fonts = []
    in_fonts = []
    in_fonts.append(parse(book, 'font_name'))
    in_fonts.append(parse(book, 'size'))
    in_fonts.append(parse(book, 'bold'))
    in_fonts.append(parse(book, 'italic'))
    for font in in_fonts:
        try:
            font_stile = font[0].childNodes[0].data
        except IndexError:
            fonts.append(None)
        else:
            fonts.append(font_stile)
    return fonts

def aligment(book):
    aligments=[]
    in_xml = []
    in_xml.append(parse(book, 'horizontal'))
    in_xml.append(parse(book, 'vertical'))
    in_xml.append(parse(book, 'wrap_text'))
    for aligment in in_xml:
        try:
            aligment_style = aligment[0].childNodes[0].data
        except IndexError:
            aligments.append(False)
        else:
            aligments.append(aligment_style)
    return aligments

def summ(SourceFile:str):
    '''Суммируем значение всх тегов tag, в записях с интексами из списка array, в файле SourcePath'''
    #try:
    #    SourceFile = modules.xml.dom.minidom.parse (SourcePath)
    #except FileNotFoundError:
    #    return 'Файл '+ SourcePath + ' не существует'
    
    books = SourceFile.getElementsByTagName("SCHET")
    id_smo = books[0].getElementsByTagName("PLAT")[0].childNodes[0].data
    nschet = books[0].getElementsByTagName("NSCHET")[0].childNodes[0].data
    data_schet = books[0].getElementsByTagName("DSCHET")[0].childNodes[0].data
    month_schet = books[0].getElementsByTagName("MONTH")[0].childNodes[0].data
    year_schet = books[0].getElementsByTagName("YEAR")[0].childNodes[0].data
    schet = [id_smo, nschet, data_schet, month_schet, year_schet]

    titles = []
    
    if SourceFile.getElementsByTagName("PLAT")[0].childNodes[0].data != '61010':
        result = [modules.Decimal('0.00')]*28

        ambuObj = books[0].getElementsByTagName("SUMMA_PF")[0]
        smpObj = books[0].getElementsByTagName("SUMMA_SMP")[0]
        result[20] += modules.Decimal(ambuObj.childNodes[0].data)
        result[22] += modules.Decimal(smpObj.childNodes[0].data)

        books = parse(SourceFile, "SLUCH")

        for book in books:
            svod = int(book.getElementsByTagName("NSVOD")[0].childNodes[0].data)
            summ = modules.Decimal(book.getElementsByTagName("SUMV")[0].childNodes[0].data)
        
            try:
                dispObj = book.getElementsByTagName("DISP")[0]
            except IndexError:
                disp='False'
            else:
                disp = dispObj.childNodes[0].data
            
            if svod//100 == 1:          #Круглосуточный стационар
                if svod-100 == int(month_schet):
                    result[0] += summ
                else:
                    result[1] += summ
                continue
            elif svod//100 == 2:        #Дневной стационар
                if svod-200 == int(month_schet):
                    result[3] += summ
                else:
                   result[4] += summ
                continue
            elif svod//100 == 3:        #Амболаторная помощь
                if svod-300 == int(month_schet):
                    result[6] += summ
                    if dispancer_adult_I.find(disp) != -1:
                        result[7] += summ
                    elif dispancer_adult_II.find(disp) != -1:
                        result[8] += summ
                    elif profosmotr_adult.find(disp) != -1:
                        result[9] += summ
                    elif dispancer_minors_difficult_I.find(disp) != -1:
                        result[10] += summ
                    elif dispancer_minors_care_I.find(disp) != -1:
                        result[11] += summ
                    elif medical_exam_minors_I.find(disp) != -1 or medical_exam_minors_II.find(disp) != -1:
                        result[12] += summ
                else:
                    result[13] += summ
                    if dispancer_adult_I.find(disp) != -1:
                        result[14] += summ
                    elif dispancer_adult_II.find(disp) != -1:
                        result[15] += summ
                    elif profosmotr_adult.find(disp) != -1:
                        result[16] += summ
                    elif dispancer_minors_difficult_I.find(disp) != -1:
                        result[17] += summ
                    elif dispancer_minors_care_I.find(disp) != -1:
                        result[18] += summ
                    elif medical_exam_minors_I.find(disp) != -1 or medical_exam_minors_II.find(disp) != -1:
                        result[19] += summ
                continue
            else:                       #Скорая медицинская помощь
                if svod-400 == int(month_schet):
                    result[22] += summ
                else:
                    result[24] += summ
    else:
        result = [modules.Decimal('0.00')]*19
        books = parse(SourceFile, "SLUCH")
        for book in books:
            svod = int(book.getElementsByTagName("NSVOD")[0].childNodes[0].data)
            summ = modules.Decimal(book.getElementsByTagName("SUMV")[0].childNodes[0].data)

            if svod//100 == 1:          #Круглосуточный стационар
                if svod-100 == int(month_schet):
                    result[1] += summ
                else:
                    result[8] += summ
                continue
            elif svod//100 == 2:        #Дневной стационар
                if svod-200 == int(month_schet):
                    result[2] += summ
                else:
                    result[9] += summ
                continue
            elif svod//100 == 3:        #Амболаторная помощь
                if svod-300 == int(month_schet):
                    result[3] += summ
                else:
                    result[10] += summ
                continue
            else:                       #Скорая медицинская помощь
                if svod-400 == int(month_schet):
                    result[4] += summ
                else:
                    result[13] += summ
    return result, schet

def settings_MO():
    file_settings = modules.xml.dom.minidom.parse ('settings.xml')
    idMO = file_settings.getElementsByTagName('id_MO')[0].childNodes[0].data
    MO = file_settings.getElementsByTagName('MO')[0].childNodes[0].data
    return idMO, MO