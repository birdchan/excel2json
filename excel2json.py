
import xlrd
import json

from argparse import ArgumentParser


def convert_col_i_2_col_ch(col_i):
    head = col_i / 26
    tail = col_i % 26
    ch = chr(ord('A')+tail)
    if (head > 0):
        return convert_col_i_2_col_ch(head-1) + ch
    else:
        return ch

def convertXls2Json(filename):

    try:
        #wb = xlrd.open_workbook(file_contents=my_file_content, use_mmap=0)
        wb = xlrd.open_workbook(filename)
    except Exception as e:
        # terminate and pass to stderr? do this later...
        raise e
        #return None

    # read the xls file content
    sheets = {}
    for sheetname in wb.sheet_names():
        sh = wb.sheet_by_name(sheetname)
        nrows, ncols = sh.nrows, sh.ncols
        my_sheet_rows = {}
        for row_i in xrange(nrows):
            row_values = sh.row_values(row_i)
            row_values_len = len(row_values)
            my_row = []
            for col_i in xrange(row_values_len):
                v = row_values[col_i]
                col_ch = convert_col_i_2_col_ch(col_i);
                my_row.append({'row': row_i+1, 'col': col_ch, 'col_i': col_i, 'val': v})
            # append to the sheet rows obj
            my_sheet_rows[row_i+1] = my_row
        # append to the overall rows obj
        sheets[sheetname] = my_sheet_rows

    # output
    return json.dumps(sheets)


if __name__ == '__main__':
    parser = ArgumentParser(description="Convert a xls file to a csv file.")
    parser.add_argument("-i", "--inputfile", required=True, dest="inputfile", help="The input xls file name.")
    parser.set_defaults(debug=False)
    args = parser.parse_args()
    
    # get values from provided args
    inputfile = args.inputfile

    # main logic
    json_str = convertXls2Json(inputfile)
    print json_str
