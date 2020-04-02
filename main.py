import modules

def stac():
    modules.svod_stac.svod_ks_ds_tfoms(source_dir, source, ks)
    print('свод кс по смо с ТФОМС')
    modules.svod_stac.svod_ks_ds_tfoms(source_dir, source, ds)
    print('свод дс по смо')
    modules.svod_stac.svod_ks_ds(source_dir, source, ks)
    print('свод кс по смо без ТФОМС')
    #modules.svod_stac.svod_ks_ds(source_dir, source, ds)
    #print('свод дс по смо без ТФОМС')
    modules.svod_stac_smo.svod_smo_ks_ds(source_dir, source, ks)
    print('свод кс по каждой смо')
    modules.svod_stac_smo.svod_smo_ks_ds(source_dir, source, ds)
    print('свод дс по каждой смо')
    modules.svod_stac_smo.svod_ks_ds_tfoms(source_dir, source, ks)
    print('свод общий кс с ТФОМС')
    modules.svod_stac_smo.svod_ks_ds_tfoms(source_dir, source, ds)
    print('свод общий дс')
    modules.svod_stac_smo.svod_ks_ds(source_dir, source, ks)
    print('свод общий кс без ТФОМС')
    #modules.svod_stac_smo.svod_ks_ds(source_dir, source, ds)
    #print('свод общий дс без ТФОМС')

def ambul():
    modules.svod.svod_amb(source_dir, source)
    print('Поликлиника')
    modules.svod_amb_smo.svod_smo_amb(source_dir, source)
    print('Поликлиника, свод по страховым')
    modules.svod_amb_smo.svod_amb_tfoms(source_dir, source)
    print('Поликлиника, свод по больнице с ТФОМС')
    modules.svod_amb_smo.svod_amb(source_dir, source)
    print('Поликлиника, свод по больнице без ТФОМС')

print('now is', modules.time.ctime())
settings = 'settings.xml'
smo = 'sp_smo.xml'

ks = 1
ds = 2

if modules.os.path.exists(settings):
    work_dir = modules.parse.value(settings, 'work_dir')
    source_dir = modules.parse.value(settings, 'source_dir')
    codeLPU = modules.parse.value(settings, 'lpuId')
else:
    print('Файл с настройками не найден')
    modules.sys.exit(0)

if work_dir.find('не содержит') != -1 or source_dir.find('не содержит') != -1 or codeLPU.find('не содержит') != -1:
    print('Файл с настройками поврежден')
    modules.sys.exit(0)

file_ok = False
bill_go = False

#Ищем файл подходящий под параметры ответа из ТФОМС
fileAnswerTFOMS = modules.files.FindAnswerTFOMS(source_dir, codeLPU)
if fileAnswerTFOMS == False:
    modules.sys.exit(0)
#Очищаем временную папку
modules.files.clear_dir(work_dir)
#Ищем и распаковывам файл с ответом
modules.files.UnZipFile(fileAnswerTFOMS, work_dir)

#распаковываем архив во временную папку
list_f = modules.os.listdir(path=str(work_dir))
for name in list_f:
    file = work_dir + name
    if name.find(codeLPU) != -1 and name.find('.zip') != -1 and name.find('_err') == -1 and name.find('_99') == -1:
        bill_go = True
        dir_out = work_dir + modules.os.path.splitext(name)[0]
        modules.os.mkdir (dir_out)
        modules.files.UnZipFile(file, dir_out)
    modules.os.remove(file)

if bill_go:
    modules.os.remove(fileAnswerTFOMS)
    
    modules.files.clearAnswer(work_dir)
    source = []
    for name_dir in modules.os.listdir( path= str (work_dir)):
        for name_file in modules.os.listdir( path= str (work_dir + name_dir)):
            file = work_dir + name_dir + modules.parse.delimiter + name_file
            source.append(modules.xml.dom.minidom.parse (file))
    for i in range (len(source)):
        result, id_smo = modules.parse.summ(source[i])
        modules.bill.bill(result, id_smo, source_dir)

    ambul()
    stac()
else:
    print('Нечего выполнять')
print('now is', modules.time.ctime())