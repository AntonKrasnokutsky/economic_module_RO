import modules

format_file = 'format.xml'
smo_file = 'smo.xml'
provder_file = 'provider.xml'
files=(provder_file, smo_file, format_file)

def preparing_of_report(SourceFile, bill_sheet, id_smo, result):
    try:
        smoObj = modules.parse.parse(SourceFile, 'smo_id')
        smo = smoObj[0].childNodes[0].data
    except IndexError:
        pass
    else:
        nodes_smo = modules.parse.parse(SourceFile, 'smo')
        for node_smo in nodes_smo:
            dataObj = modules.parse.parse(node_smo, 'smo_id')
            for smoid in dataObj:
                data = smoid.childNodes[0].data
                if data.find(id_smo[0]) != -1:
                    nodes_smo = node_smo

    modules.outxlsx.set_format(bill_sheet, nodes_smo)
    #ширина ячеек
    modules.outxlsx.set_width(bill_sheet, nodes_smo)
    #высота ячеек
    modules.outxlsx.set_height(bill_sheet, nodes_smo)
    modules.outxlsx.set_months(bill_sheet, nodes_smo, id_smo, result)
    
    try:
        cell_id_Obj = modules.parse.parse(SourceFile, 'cell_id')
        cell_id = cell_id_Obj[0].childNodes[0].data
    except IndexError:
        pass
    else:
        cell_result_Obj = modules.parse.parse(nodes_smo, 'result')
        for cell_result in cell_result_Obj:
            cell_id_Obj = modules.parse.parse(cell_result, 'cell_id')
            cell_id = cell_id_Obj[0].childNodes[0].data
            result_id_Obj = modules.parse.parse(cell_result, 'result_id')
            result_id = int(result_id_Obj[0].childNodes[0].data)
            bill_sheet[cell_id].value = result[result_id]

def oms_format(bill_sheet_out, result, id_smo):
    cell_calc(result, id_smo[0])
    for filename in files:
        try:
            SourceFile = modules.xml.dom.minidom.parse (filename)
        except FileNotFoundError:
            print ('Файл ' + SourcePath + ' не существует')
        SourceFile.normalize()
        preparing_of_report(SourceFile, bill_sheet_out, id_smo, result)

def cell_calc(result, id_smo):
    if id_smo != '61010':
        result[2] = result[0] + result[1]
        result[5] = result[3] + result[4]
        result[21] = result[6] + result[13] + result[20]
        result[26] = result[22] + result[24]
        result[27] = result[2] + result[5] + result[21] + result[26]
    else:
        result[0] = result[1] + result[2] + result[3] + result[4]
        result[5] = result[6]
        result[7] = result[8] + result[9] + result[10] + result[11]
        result[12] = result[13]
        result[15] = result[1] + result[8]
        result[16] = result[2] + result[9]
        result[17] = result[3] + result[10]
        result[18] = result[4] + result[11]
        result[14] = result[15] + result[16] + result[17] + result[18]

def bill(result, id_smo, work_dir):
    bill_out_file = modules.openpyxl.Workbook ()
    bill_out_sheet = bill_out_file.active
    bill_out_sheet.title = 'Счет'
    oms_format(bill_out_sheet, result, id_smo)
    bill_out_file.save(work_dir + modules.parse.delimiter + 'Re_' + id_smo[1] + '.xlsx')
    bill_out_file.close
    print('Подготовка счета:' , id_smo[0], 'Ок')