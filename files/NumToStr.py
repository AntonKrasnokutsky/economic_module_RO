#Функция подписи копеек правильно
def koptxt(kop):
    #       0         1         2         3         4         5
    txt=['копеек','копейка','копейки','копейки','копейки','копеек','копеек','копеек','копеек','копеек']
    if kop<20 and kop>10:
        return 'копеек'
    else:
        t2=int(kop/10)*10
        tm=kop-t2
        return txt[tm]
