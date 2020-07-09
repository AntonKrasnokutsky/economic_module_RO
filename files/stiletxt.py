import openpyxl

def font(name, size ,bold=False, italic=False):
    return openpyxl.styles.Font(name=str(name),
                                size=float(size),
                                bold=int(bold),
                                italic=int(italic))

border_none=('00000000','00000000','00000000','00000000')

def borders(border=border_none):
    bor=openpyxl.styles.Border()
    if border[0]!=None:
        bor.left=openpyxl.styles.Side(border_style='thin', color=border[0])
    if border[1]!=None:
        bor.right=right=openpyxl.styles.Side(border_style='thin', color=border[1])
    if border[2]!=None:
        bor.top=top=openpyxl.styles.Side(border_style='thin',color=border[2])
    if border[3]!=None:
        bor.bottom=openpyxl.styles.Side(border_style='thin',color=border[3])

    return bor

def alig(ali_h='general', ali_v='top', wrap=False):
    return openpyxl.styles.Alignment(horizontal=str(ali_h),
                                     vertical=str(ali_v),
                                     wrap_text=int(wrap))