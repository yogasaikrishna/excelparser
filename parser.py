import sys
import xlrd


def parse():
    file = sys.argv[1] if len(sys.argv) > 1 else None
    sheet_number = int(sys.argv[2]) if len(sys.argv) > 2 else 0
    if not file:
        print 'Please provide source file path'
        return 1
    try:
        book = xlrd.open_workbook(file)
        sheet = book.sheet_by_index(sheet_number)
        print "No. of worksheets: {0}".format(book.nsheets)
        print "Worksheet names: {0}".format(book.sheet_names())
        print "{0} {1} {2}".format(sheet.name, sheet.nrows, sheet.ncols)
    except Exception as error:
        print 'Exception - ' + str(error.message)

parse()
