"""
File names writing in excel file
If you want to write them vertically, you need switch the "i" and "0" at the line 35
"""

import os
import xlwt
def main():
    wb = xlwt.Workbook()
    ws = wb.add_sheet('barisx')
    i = 0
    path  = 'C:\\Users\\<YourPCName>\\Desktop\\<FilesFolder>\'
    characters = "/1234567890><(')¢*&©^%°£$—®»£;`?¬\"|!]~,“”@é}[.:=’‘" # we don't want to see these characters
    for filename in os.listdir(path):
        src = filename
        for c in characters:
            src = src.translate({ord(c): None})
        dst = src.replace( '-', ' ' ) # sometimes "-" should change with blank 
        if 'jpg' in dst:
            dst = dst.replace( 'jpg', '' ) # any file type need to destroy
        elif 'jpeg' in dst:
            dst = dst.replace( 'jpeg', '' ) # you can add one or more like this
        elif 'Png' in dst:
            dst = dst.replace( 'Png', '' )
        elif 'Jpg' in dst:
            dst = dst.replace( 'Jpg', '' )
        elif 'JPG' in dst:
            dst = dst.replace( 'JPG', '' )
        elif 'PDF' in dst:
            dst = dst.replace( 'PDF', '' )
        else:
            dst = dst[:-3]
        print("Writing name:"+dst+"\nCount:"+str(i+1))

        ws.write(i,0, dst)
        i = i + 1
    wb.save('filenametoexcel.xls')

if __name__ == '__main__':

    # Calling main() function
    main()
