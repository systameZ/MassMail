
import xlrd

class read_xlsx(object):
    def __call__(self,path):
        workbook=xlrd.open_workbook(path)
        booksheet=workbook.sheet_by_index(0)
        ret_dict={}
        try:
                for row in range(booksheet.nrows):
                    ret_dict[str(booksheet.cell(row,0).value)]=[str(booksheet.cell(row,1).value),str(booksheet.cell(row,2).value)]
        except IndexError as e:
            for row in range(booksheet.nrows):
                ret_dict[str(booksheet.cell(row, 0).value)] = str(booksheet.cell(row, 1).value)
        except Exception as ee:
            return 'Error : '+ str(ee)
        return ret_dict

if __name__ =='__main__':
    red=read_xlsx()
    print(red('D:\\111.xlsx'))
    for i in red('D:\\111.xlsx'):
        print(i)

