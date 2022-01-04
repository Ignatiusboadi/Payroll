#Module imports
from numpy import zeros
from random import choice, choices
from string import ascii_lowercase, ascii_uppercase, digits
import xlsxwriter as xlsx

payroll = xlsx.Workbook('PDA PAYROLL.xlsx') #Name of Excel File

#Cell formats
empty_format = '_(* #,##0.00_);_(* (#,##0.00);_(* ""_);_(@_)'
bold = payroll.add_format({'bold':True})
bold_center = payroll.add_format({'bold':True, 'align':'center'})
center = payroll.add_format({'align':'center'})
locked = payroll.add_format({'locked':True})
_25 = payroll.add_format({'bold':True, 'align':'center','font_name':'Arial Black', 'font_size':20,
                     'border':2, 'bg_color':'#D2691E'})
_17 = payroll.add_format({'bold':True, 'align':'center', 'bg_color':'#D2691E', 'font_name':'Cambria',
                     'font_size':17, 'border':2})
_15_270 = payroll.add_format({'bold':True, 'align':'center', 'bg_color':'#D2691E',
                              'font_name':'Cambria', 'font_size':15, 'border':2, 'valign':'Top'})
_15_270.set_rotation(270)
_12 = payroll.add_format({'border':1, 'font_size':13, 'font_name':'Bookman Old Style',
                          'bg_color':'silver', 'num_format':empty_format})
_12_custom = payroll.add_format({'border':1, 'font_size':13, 'font_name':'Cambria',
                                 'text_wrap':True, 'bg_color':'#D2691E', 'num_format':empty_format})
_12_txt = payroll.add_format({'border':1, 'font_size':13, 'font_name':'Bookman Old Style',
                              'text_wrap':True, 'bg_color':'#D2691E', 'num_format':'_(* ""_);_(@_)'})
_13_gray = payroll.add_format({'border':2, 'font_size':13, 'font_name':'Cambria', 'text_wrap':True,
                              'bg_color':'gray', 'align':'center','num_format':empty_format,
                               'bold':True})
_12_gold = payroll.add_format({'border':1, 'font_size':13, 'font_name':'Bookman Old Style', 
                               'bg_color':'#FFD700', 'num_format':empty_format})
black = payroll.add_format({'bg_color':'black'})

#initial values
start_year = 2008
end_year = 2081
staff_num = 150

months = ['JAN', 'FEB', 'MAR', 'APR', 'MAY', 'JUN', 'JUL', 'AUG', 'SEP', 'OCT', 'NOV', 'DEC']
months_names = ['JANUARY', 'FEBRUARY', 'MARCH', 'APRIL', 'MAY', 'JUNE', 'JULY', 'AUGUST',\
                'SEPTEMBER', 'OCTOBER', 'NOVEMBER', 'DECEMBER']

heads = ['NAME OF EMPLOYEE', 'POSITION', 'MOBILE NO.', 'BASIC SALARY', 'ALLOWANCES', \
         'GROSS CONSOLIDATED', 'EXCESS BONUS', 'TOTAL GROSS CONSOLIDATED', 'SSF - 5.5% ON BASIC',\
         'TAXABLE INCOME', 'INCOME TAX DEDUCTED', 'LESS PROVIDENT FUND - 5% ON BASIC',\
         'NET SALARY', 'LESS WELFARE DUES', 'LESS WELFARE LOAN', 'LESS PROVIDENT FUND LOAN',\
         'LESS RENT LOAN', 'NET TAKE HOME', '', 'EMPLOYERS\' 13% CONTRIBUTION',\
         'EMPLOYERS\' 5% PROVIDENT FUND CONTRIBUTION', 'COMMENTS']

heads = [f'{head} (GH₵)' if head not in ['', 'COMMENTS', 'NAME OF EMPLOYEE', 'POSITION', 'MOBILE NO.']\
         else head for head in heads]

f_rows = ['PDA STAFF PAYROLL FOR ', months] # first rows common to all sheets
pay_heading = 'PDA COMPREHENSIVE PAYROLL'   # heading common to all sheets
f_rows_num = len(f_rows) + 1                #number of first rows
n_cols = len(heads)                         #number of first columns
xl_cols = [alpha + beta for alpha in [''] + list(ascii_uppercase) for beta in ascii_uppercase]         #Excel column names
pay_sheets = {year:payroll.add_worksheet(f'{year} PAYROLL') for year in range(start_year, end_year+1)} #Payroll Sheets

for year in range(start_year, end_year + 1):
    sheet = pay_sheets[year]
    sheet.set_row(0, 25)
    sheet.set_row(1, 15)
    sheet.set_row(2, 16)
    sheet.freeze_panes(5, 0)
    for row in range(50 + f_rows_num + 3, staff_num//2 + f_rows_num + 3 + 6, 10):
        for i in range(9):
            sheet.set_row(row + i, None, None, {'hidden':True, 'level':1})
    for row in range(staff_num//2 + f_rows_num + 3 + 7, staff_num + f_rows_num, 33):
        for i in range(32):
            sheet.set_row(row + i, None, None, {'hidden':True, 'level':1})
#     Top three rows merging and inserting text and pictures
    sheet.insert_image('A1', 'pda logo.png', {'x_scale':0.55, 'y_scale':0.55})
    sheet.merge_range(0, 0, 1, -1 + (n_cols + 1) * 12, f'PDA STAFF PAYROLL {year}', _25)
    sheet.merge_range(2, 0, 2, -1 + (n_cols + 1) * 12, '', _13_gray)
    
    ind = 0
    for index, month in enumerate(months):
        sheet.data_validation(f_rows_num + 2, ind + 3, staff_num + f_rows_num + 1, ind + 4,
        {'validate':'decimal', 'criteria':'>=', 'minimum':0,
         'input_message':'Enter a figure greater than or equal to 0'})
        sheet.data_validation(f_rows_num + 2, ind + 6, staff_num + f_rows_num + 1, ind + 6,
        {'validate':'decimal', 'criteria':'>=', 'minimum':0,
         'input_message':'Enter a figure greater than or equal to 0'})
        sheet.data_validation(f_rows_num + 2, ind + 13, staff_num + f_rows_num + 1, ind + 16,
        {'validate':'decimal', 'criteria':'>=', 'minimum':0,
         'input_message':'Enter a figure greater than or equal to 0'})
        sheet.set_column(ind, ind, 50, None, {'hidden':True, 'level':1})
        sheet.set_column(ind + 1, ind + 1, 25, None, {'hidden':True, 'level':1})
        sheet.set_column(ind + 2, ind + n_cols - 1, 20, None, {'hidden':True, 'level':1})
        sheet.set_column(ind + n_cols, ind + n_cols, 7)
        sheet.merge_range(f_rows_num, ind, f_rows_num, ind + n_cols - 1, 
                          f'{pay_heading} {months_names[index]} {year}' , _17)
        sheet.merge_range(f_rows_num, ind + n_cols, f_rows_num + 1, ind + n_cols, '', _15_270)
        sheet.merge_range(f_rows_num + 2, ind + n_cols, f_rows_num + staff_num + 1, ind + n_cols,
                          f'END OF {months_names[index]}', _15_270)
        sheet.merge_range(staff_num + f_rows_num + 2, ind , staff_num + f_rows_num + 2, ind + 2,
                      'TOTALS', _13_gray)
        sheet.write_row(f'{xl_cols[ind]}{f_rows_num + 2}:{xl_cols[ind + n_cols - 1]}{f_rows_num + 2}',
                        heads, _13_gray)
        staff_names = ['' for j in range(staff_num)]
        row_range = range(f_rows_num + 3, f_rows_num + staff_num + 3)
        gross_cons = [f'={xl_cols[ind + 3]}{row}+{xl_cols[ind + 4]}{row}' 
                          for row in row_range]
        total_gross_cons = [f'={xl_cols[ind + 5]}{row}+{xl_cols[ind + 6]}{row}' 
                      for row in row_range]
        ssnit = [f'={xl_cols[ind + 3]}{row}*5.5%' 
                      for row in row_range]
        tax_income = [f'={xl_cols[ind + 7]}{row}-{xl_cols[ind + 8]}{row}' 
                      for row in row_range]
        cell = xl_cols[ind + 9]
        tax = ['=IF({}{} <= 319, 0, \
               IF({}{} <= 419, 5% * ({}{} - 319), \
               IF({}{} <= 539, 5 + (10% * ({}{} - 419)),\
               IF({}{} <= 3539, 17 + (17.5% * ({}{} - 539)),\
               IF({}{} <= 20000, 542 + (25%* ({}{} - 3539)),\
               4657.25 + 30%*({}{}-20000))))))'.format(
            cell, row, cell, row, cell, row, cell, row, cell, row, cell, row, cell, row, cell, row,
            cell, row, cell, row) for row in row_range]
        prov_fund = [f"={xl_cols[ind + 3]}{row}*5%" for row in row_range]
        net_salary = [
            f'={xl_cols[ind + 9]}{row} - {xl_cols[ind + 10]}{row} - {xl_cols[ind + 11]}{row}'\
            for row in row_range]
        take_home = [f'={xl_cols[ind + 12]}{row} - SUM({xl_cols[ind + 13]}{row}:{xl_cols[ind + 16]}{row})'\
                     for row in row_range]
        emp_cont = [f"={xl_cols[ind + 3]}{row}*13%" for row in row_range]
        emp_prov_cont = [f"={xl_cols[ind + 3]}{row}*5%" for row in row_range]
        totals = [
            f'=SUM({xl_cols[col_index]}{f_rows_num + 2}:{xl_cols[col_index]}{f_rows_num + staff_num + 2})'
        for col_index in range(ind + 3, n_cols + ind + 1)]
        cols_sel = [0, 1, 2, 3, 4, 6, 13, 14, 15, 16, 18, 21]
        fmts = [_12_txt if col_sel == 2 else _12_custom if col_sel in [0, 1] else _12_gold \
                if col_sel in cols_sel else _12 for col_sel in range(n_cols)]
        sheet.write_row(staff_num + f_rows_num + 2, ind + 3, totals, _13_gray)
        if index == 0:
            if year == start_year:
                formulas = [gross_cons, total_gross_cons, ssnit,\
                        tax_income, tax, prov_fund, net_salary, take_home, emp_cont, emp_prov_cont]
                cols_sel = [0, 1, 2, 3, 4, 6, 13, 14, 15, 16, 18, 21]
                j = 0
                for i in range(n_cols):
                    if i in cols_sel:
                        sheet.write_column(f_rows_num + 2, ind + i, staff_names, fmts[i])
                    else:
                        sheet.write_column(f_rows_num + 2, ind + i, formulas[j], fmts[i])
                        j += 1
            else:
                prev_names = [f"='{year - 1} PAYROLL'!{xl_cols[prev_col]}{row}" for row in row_range]
                prev_post = [f"='{year - 1} PAYROLL'!{xl_cols[prev_col + 1]}{row}" for row in row_range]
                prev_mob = [f"='{year - 1} PAYROLL'!{xl_cols[prev_col + 2]}{row}" for row in row_range]
                ini = [prev_names, prev_post, prev_mob]
                formulas = [gross_cons, total_gross_cons, ssnit,\
                        tax_income, tax, prov_fund, net_salary, take_home, emp_cont, emp_prov_cont]
                cols_sel = [0, 1, 2, 3, 4, 6, 13, 14, 15, 16, 18, 21]
                j = 0
                for i in range(n_cols):
                    if i in [0, 1, 2]:
                        sheet.write_column(f_rows_num + 2, ind + i, ini[i], fmts[i])
                    elif i in cols_sel:
                        sheet.write_column(f_rows_num + 2, ind + i, staff_names, fmts[i])
                    else:
                        sheet.write_column(f_rows_num + 2, ind + i, formulas[j], fmts[i])
                        j += 1
        else:
            name_entry = [f'={xl_cols[ind - n_cols - 1]}{row}' for row in row_range]
            position_entry = [f'={xl_cols[ind - n_cols]}{row}' for row in row_range]
            mobile_entry = [f'={xl_cols[ind - n_cols + 1]}{row}' for row in row_range]
            formulas = [name_entry, position_entry, mobile_entry, gross_cons, total_gross_cons,\
                        ssnit, tax_income, tax, prov_fund, net_salary, take_home, emp_cont,\
                        emp_prov_cont]
            cols_sel = [3, 4, 6, 13, 14, 15, 16, 18, 21]
            j = 0
            for i in range(n_cols):
                if i in cols_sel:
                    sheet.write_column(f_rows_num + 2, ind + i, staff_names, fmts[i])
                else:
                    sheet.write_column(f_rows_num + 2, ind + i, formulas[j], fmts[i])
                    j += 1
        sheet.set_column(ind + n_cols - 4, n_cols + ind - 4, 3, None, {'level': 1, 'hidden': True})
        sheet.merge_range(f_rows_num + 1, ind + n_cols - 4, f_rows_num + staff_num + 2, 
                          ind + n_cols - 4, '', black)
        ind += (n_cols + 1)
        if month == 'DEC':
            prev_col = ind - n_cols - 1
    
payroll.close()

