import os
import shutil
import zipfile
import parse

def clear_dir(path):
    """ Очистка временной папаки от всего содержимого
На входе необходимо указать путь к папке для очистки"""
    for name in os.listdir (path = str (path)):
        file = path + name
        try:
            os.remove(file)
        except IsADirectoryError:
            shutil.rmtree(file)
        except PermissionError:
            shutil.rmtree(file)

def FindAnswerTFOMS(path, kodLPU):
    """Ищем ответный фал из ТФОМС в папке path и возвращаем путь если нашли
вторым параметром указывается код ЛПУ из справочника ТФОМС"""
    for name in os.listdir (path = str (path)):
        if name.find(kodLPU) != -1 and name.find('.zip') != -1:
            return path + name
    return False

def UnZipFile(SourceFile, DistenationPath):
    '''SourceFile - путь к архиву, DistenationPath - путь к целевой папке
 Распаковка при условии наличия исходного файла'''
    if SourceFile != False:
        tfom_zip = zipfile.ZipFile (SourceFile)
        tfom_zip.extractall(DistenationPath)
        tfom_zip.close

def clearAnswer(SourcePath):
    '''Удалям ненужные файлы, входнеой параметр SourcePath указывает на рабочую папку
Удаялем все фалы не содержащие нужную информацию, по ТС нужные файлы начинаются с "HM"'''
    for name_dir in os.listdir( path= str (SourcePath)):
        for name_file in os.listdir( path= str (SourcePath + name_dir)):
            file = SourcePath + name_dir + parse.delimiter + name_file
            if name_file.find('HM') == -1:
                os.remove(file)