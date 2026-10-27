# Expense Report Generator
# Written by Weston Matthews (2020)

# Takes a collection of excel expense reports, exported by QuickBooks, and uses them to create an excel spreadsheet which lists expenses by location and by month
# Compares this year's expenses to last year's expenses
# Also produces a report on summary Year To Date expenses, including Dollar Change and % Change
# Produces a Month comparison report if the latest month in the report is not December.
# The list of locations/books is read from config.json (see config.example.json).

import openpyxl as xl

import json

import sys

import __main__

import easygui as gui

from datetime import datetime

from copy import copy

import os

def detect_latest_month(sheet):
    # Takes in a sheet object and reads the top line, looking to find the last month in the list
    max_col = sheet.max_column
    latest_month = 0
    for x in range(1, max_col+1):
        col_letter = sheet.cell(column=x, row=1).column_letter
        cell_data = sheet[col_letter+'1'].value
        if cell_data != None:
            #print(cell_data)
            if "Jan" in cell_data:
                latest_month = 1
            elif "Feb" in cell_data:
                latest_month = 2
            elif "Mar" in cell_data:
                latest_month = 3
            elif "Apr" in cell_data:
                latest_month = 4
            elif "May" in cell_data:
                latest_month = 5
            elif "Jun" in cell_data:
                latest_month = 6
            elif "Jul" in cell_data:
                latest_month = 7
            elif "Aug" in cell_data:
                latest_month = 8
            elif "Sep" in cell_data:
                latest_month = 9
            elif "Oct" in cell_data:
                latest_month = 10
            elif "Nov" in cell_data:
                latest_month = 11
            elif "Dec" in cell_data:
                latest_month = 12
            elif cell_data == "TOTAL":
                return latest_month
    if latest_month == 0:
        return -1

def copy_cell(write_cell, trans_cell):
    # Copies every parameter of one cell to another
    # Can be between worksheets or workbooks
    # Value
    trans_value = trans_cell.value
    # Font
    trans_font =  copy(trans_cell.font)
    # Fill
    trans_fill = copy(trans_cell.fill)
    # Border
    trans_border = copy(trans_cell.border)
    # Alignment
    trans_align = copy(trans_cell.alignment)
    # Number Format
    trans_format = trans_cell.number_format
    
    #print("Cell value: "+str(trans_value))
    #print("Cell font: "+str(trans_font))
    #print("Cell fill: "+str(trans_fill))
    #print("Cell align: "+str(trans_align))
    #print("Cell format: "+str(trans_format))
    
    write_cell.value = trans_value
    write_cell.font = trans_font
    write_cell.fill = trans_fill
    write_cell.border = trans_border
    write_cell.alignment = trans_align
    write_cell.number_format = trans_format

def find_month_total(sheet, month):
    # Find the total for the month for the sheet
    # Month must be three characters long
    month = month[:3]
    for col in range(1, sheet.max_column+1):
        col_letter = sheet.cell(column=col, row=1).column_letter
        cell_data = sheet[col_letter+'1'].value
        if cell_data != None and month in cell_data:
            val = sheet[col_letter+str(sheet.max_row)].value
            row = sheet.max_row
            while True:
                if val == None:
                    row -= 1
                    val = sheet[col_letter+str(row)].value
                else:
                    return val
    return "ERR"

# Maximum number of locations the summary layout has room for (rows 4-15 / 26-37)
MAX_LOCATIONS = 12
CONFIG_NAME = "config.json"
EXAMPLE_CONFIG_NAME = "config.example.json"

def app_dir():
    # Folder containing the script, or the .exe when built with PyInstaller
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))

def find_config():
    # Looks for config.json in the working directory, then next to the script/exe.
    # Falls back to config.example.json so the tool can be tried out of the box.
    for folder in (os.getcwd(), app_dir()):
        candidate = os.path.join(folder, CONFIG_NAME)
        if os.path.exists(candidate):
            return candidate
    for folder in (os.getcwd(), app_dir()):
        candidate = os.path.join(folder, EXAMPLE_CONFIG_NAME)
        if os.path.exists(candidate):
            print("No "+CONFIG_NAME+" found, using "+EXAMPLE_CONFIG_NAME)
            return candidate
    return None

def load_locations(config_path):
    # Returns a list of (name, code) tuples, in summary-table order.
    # name: what the user picks from the list; code: short label used for sheet names and the summary table
    with open(config_path, "r", encoding="utf-8") as f:
        config = json.load(f)
    locations = []
    for entry in config.get("locations", []):
        name = str(entry["name"])
        code = str(entry.get("code", name))
        locations.append((name, code))
    if len(locations) == 0:
        raise ValueError("No locations defined in "+config_path)
    if len(locations) > MAX_LOCATIONS:
        raise ValueError("Too many locations in "+config_path+" (max "+str(MAX_LOCATIONS)+")")
    return locations

def find_total(sheet):
    col_letter = sheet.cell(column=sheet.max_column, row=sheet.max_row).column_letter
    return sheet[col_letter+str(sheet.max_row)]
    

def main():
    # Collect the expense reports for each location, for this year and for last year
    # Read-in the information in each expense report, organize it, and store it in memory
    while True:
        msg = "What year is this report for?"
        title = "Expense Report Generator v1.0"
        cur_year = gui.enterbox(msg, title)
        if cur_year == None:
            return 0
        else:
            try:
                cur_year = int(cur_year)
                if cur_year < 1000:
                    msg = "Year too low! You entered a year before 1000 AD!"
                    gui.msgbox(msg, title)
                else:
                    break
            except:
                msg = "Invalid year! You must enter a number!"
                gui.msgbox(msg, title)
    #print(cur_year)
    #cur_year = datetime.now().year
    last_year = cur_year-1
    # Create a new excel spreadsheet workbook
    wb = xl.Workbook()
    # Create a summary page, and add each location's expense report to the workbook
    ## Creating summary page
    ws_sum = wb.active
    ws_sum.title = "Summary"
    ## Adding previous year reports for each location
    msg = "Which locations?"
    title = "Expense Report Generator v1.0"
    config_path = find_config()
    if config_path == None:
        gui.msgbox("Could not find "+CONFIG_NAME+". Copy "+EXAMPLE_CONFIG_NAME+" to "+CONFIG_NAME+" and edit it.", title)
        return 0
    try:
        config_locations = load_locations(config_path)
    except (OSError, ValueError, KeyError) as err:
        gui.msgbox("Could not read "+config_path+":\n"+str(err), title)
        return 0
    choices = [name for (name, code) in config_locations]
    loc_codes = dict(config_locations)
    # easygui's multchoicebox needs at least two choices
    if len(choices) == 1:
        locations = choices
    else:
        locations = gui.multchoicebox(msg, title, choices)
    #print(locations)
    if locations == [] or locations == None:
        return 0
    else:
        paths = []
        for location in locations:
            msg = "Please select the expense report for "+location+" for "+str(cur_year)
            paths.append(gui.fileopenbox(msg, title, filetypes=["*.xlsx", "*.xls"]))
            msg = "Please select the expense report for "+location+" for "+str(last_year)
            paths.append(gui.fileopenbox(msg, title, filetypes=["*.xlsx", "*.xls"]))
        #print(paths)
    ## If file paths for report spreadhseets does not exist, or if None, skip over it
    
    #latest_month = 10 # Latest month
    latest_months = []
    sheet_names = []
    for location in locations:
        code = loc_codes[location]
        sheet_names.append(code+" "+str(cur_year)[2:])
        sheet_names.append(code+" "+str(last_year)[2:])
    #print(sheet_names)
    for x in range(0, len(paths)):
        path = paths[x]
        sheet_name = sheet_names[x]
        ## Copy the information by location and by year into new sheets
        ## Order is the order of the locations in the config file
        ## Year order: this year, last year
        # Create a new sheet, name it correctly, and transfer the data over from each expense report
        if path != None and path != "":
            print("Copying data from '"+path+"' into sheet '"+sheet_name+"'.")
            trans_wb = xl.load_workbook(filename=path, data_only=True)
            trans_ws = trans_wb.active
            max_col = trans_ws.max_column
            max_row = trans_ws.max_row
            cur_ws = wb.create_sheet(title=sheet_name)
            print("Reading...")
            for row in range(1, max_row+1):
                for col in range(1, max_col+1):
                    trans_cell = trans_ws.cell(column=col, row=row)
                    write_cell = cur_ws.cell(column=col, row=row)
                    col_letter = write_cell.column_letter
                    copy_cell(write_cell, trans_cell)
            for idx, rd in trans_ws.row_dimensions.items():
                cur_ws.row_dimensions[idx] = copy(rd)
            for idy, cd in trans_ws.column_dimensions.items():
                cur_ws.column_dimensions[idy] = copy(cd)
            cur_ws.freeze_panes = trans_ws.freeze_panes
            print("Done reading!")
    # Create the data tables, enter the information, and link the report cells together with function cells
    ## Creating data frames in "Summary"
    
    ## Current year data frame
    ws_sum.column_dimensions['A'].width = 1.4
    ws_sum.column_dimensions['B'].width = 7.0
    ws_sum.column_dimensions['C'].width = 1.4
    ws_sum.row_dimensions[2].height = 22.50
    thin = xl.styles.Side(border_style="thin", color='000000')
    thick = xl.styles.Side(border_style="thick", color='000000')
    ws_sum['D1'] = cur_year
    ws_sum['D1'].font = xl.styles.Font(name='Times New Roman', size=11, b=True)
    ws_sum['D1'].fill = xl.styles.PatternFill("solid", start_color='ffff00', end_color='ffff00')
    ws_sum['D1'].alignment = xl.styles.Alignment(horizontal="center", vertical="center")
    ws_sum['D1'].border = xl.styles.Border(top=thin, left=thin, right=thin, bottom=thin)
    mon = ["Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec", "Average/Location"]
    mon_ls = ['D2', 'E2', 'F2', 'G2', 'H2', 'I2', 'J2', 'K2', 'L2', 'M2', 'N2', 'O2', 'P2']
    
    for sheet in wb:
        if sheet.title != "Summary" and str(cur_year)[2:] in sheet.title:
            latest_months.append(detect_latest_month(sheet))
    latest_month = min(latest_months)
    cur_latest_mon = latest_month
    
    for x in range(0, len(mon_ls)):
        ws_sum[mon_ls[x]] = mon[x]
        ws_sum[mon_ls[x]].font = xl.styles.Font(name='Times New Roman', size=11, color='FFFFFF', b=True)
        ws_sum[mon_ls[x]].fill = xl.styles.PatternFill("solid", start_color='5b9ad5', end_color='5b9ad5')
        ws_sum[mon_ls[x]].alignment = xl.styles.Alignment(horizontal="center", vertical="center")
        ws_sum[mon_ls[x]].border = xl.styles.Border(top=thin, left=thin, right=thin, bottom=thin)
        ws_sum.column_dimensions[mon_ls[x][0]].width = 14.00
    ws_sum.column_dimensions['P'].width = 20.00
    if latest_month < 12:
        if latest_month == 1:
            ws_sum['Q2'] = "Total for the month"
        else:
            ws_sum['Q2'] = "Total for the "+str(latest_month)+" months"
        ws_sum['Q2'].font = xl.styles.Font(name='Times New Roman', size=11, color='FFFFFF', b=True)
        ws_sum['Q2'].fill = xl.styles.PatternFill("solid", start_color='5b9ad5', end_color='5b9ad5')
        ws_sum['Q2'].alignment = xl.styles.Alignment(horizontal="center", vertical="center")
        ws_sum['Q2'].border = xl.styles.Border(top=thin, left=thin, right=thin, bottom=thin)
        ws_sum.column_dimensions['Q'].width = 22.00
    else:
        ws_sum['Q2'] = "Total"
        ws_sum['Q2'].font = xl.styles.Font(name='Times New Roman', size=11, color='FFFFFF', b=True)
        ws_sum['Q2'].fill = xl.styles.PatternFill("solid", start_color='5b9ad5', end_color='5b9ad5')
        ws_sum['Q2'].alignment = xl.styles.Alignment(horizontal="center", vertical="center")
        ws_sum['Q2'].border = xl.styles.Border(top=thin, left=thin, right=thin, bottom=thin)
        ws_sum.column_dimensions['Q'].width = 22.00
    
    locs = [code for (name, code) in config_locations]
    for x in range(0, len(locs)):
        ws_sum['B'+str(x+4)] = locs[x]
        ws_sum['B'+str(x+4)].font = xl.styles.Font(name='Calibri', size=11, b=True)
        ws_sum['B'+str(x+4)].fill = xl.styles.PatternFill("solid", start_color='ffff00', end_color='ffff00')
        ws_sum['B'+str(x+4)].alignment = xl.styles.Alignment(horizontal="center", vertical="center")
    ws_sum.column_dimensions['B'].width = 7.00
    
    ws_sum['H20'].font = xl.styles.Font(name='Calibri', size=11)
    ws_sum['H20'].fill = xl.styles.PatternFill("solid", start_color='ffff00', end_color='ffff00')
    ws_sum['H20'].border = xl.styles.Border(top=thick, left=thick, right=thick, bottom=thick)
    ws_sum['H21'] = "Running Average"
    ws_sum['H21'].font = xl.styles.Font(name='Times New Roman', size=11, b=True)
    ws_sum['H21'].alignment = xl.styles.Alignment(horizontal="center", vertical="center")
    
    rows = [3,5,7,9,11,13,15,17]
    cols = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q']
    for x in cols:
        for y in rows:
            ws_sum[x+str(y)].fill = xl.styles.PatternFill("solid", start_color='ddebf7', end_color='ddebf7')
    
    cell_range = ws_sum['D2':'Q18']
    for row in cell_range:
        for cell in row:
            cell.border = xl.styles.Border(top=thin, left=thin, right=thin, bottom=thin)
            
    ## Last year data frame
    latest_months = []
    
    ws_sum.row_dimensions[24].height = 22.50
    ws_sum['D23'] = last_year
    ws_sum['D23'].font = xl.styles.Font(name='Times New Roman', size=11, b=True)
    ws_sum['D23'].fill = xl.styles.PatternFill("solid", start_color='ffff00', end_color='ffff00')
    ws_sum['D23'].alignment = xl.styles.Alignment(horizontal="center", vertical="center")
    ws_sum['D23'].border = xl.styles.Border(top=thin, left=thin, right=thin, bottom=thin)
    ws_sum.row_dimensions[24].height = 22.50
    
    mon_ls = ['D24', 'E24', 'F24', 'G24', 'H24', 'I24', 'J24', 'K24', 'L24', 'M24', 'N24', 'O24', 'P24']
    for x in range(0, len(mon_ls)):
        ws_sum[mon_ls[x]] = mon[x]
        ws_sum[mon_ls[x]].font = xl.styles.Font(name='Times New Roman', size=11, color='FFFFFF', b=True)
        ws_sum[mon_ls[x]].fill = xl.styles.PatternFill("solid", start_color='5b9ad5', end_color='5b9ad5')
        ws_sum[mon_ls[x]].alignment = xl.styles.Alignment(horizontal="center", vertical="center")
        ws_sum[mon_ls[x]].border = xl.styles.Border(top=thin, left=thin, right=thin, bottom=thin)
        ws_sum.column_dimensions[mon_ls[x][0]].width = 14.00
    ws_sum.column_dimensions['P'].width = 20.00
    
    for sheet in wb:
        if sheet.title != "Summary" and str(last_year)[2:] in sheet.title:
            latest_months.append(detect_latest_month(sheet))
    latest_month = min(latest_months)
    last_latest_mon = latest_month
    
    if latest_month < 12:
        if latest_month == 1:
            ws_sum['Q24'] = "Total for the month"
        else:
            ws_sum['Q24'] = "Total for the "+str(latest_month)+" months"
        ws_sum['Q24'].font = xl.styles.Font(name='Times New Roman', size=11, color='FFFFFF', b=True)
        ws_sum['Q24'].fill = xl.styles.PatternFill("solid", start_color='5b9ad5', end_color='5b9ad5')
        ws_sum['Q24'].alignment = xl.styles.Alignment(horizontal="center", vertical="center")
        ws_sum['Q24'].border = xl.styles.Border(top=thin, left=thin, right=thin, bottom=thin)
        ws_sum.column_dimensions['Q'].width = 22.00
    else:
        ws_sum['Q24'] = "Total"
        ws_sum['Q24'].font = xl.styles.Font(name='Times New Roman', size=11, color='FFFFFF', b=True)
        ws_sum['Q24'].fill = xl.styles.PatternFill("solid", start_color='5b9ad5', end_color='5b9ad5')
        ws_sum['Q24'].alignment = xl.styles.Alignment(horizontal="center", vertical="center")
        ws_sum['Q24'].border = xl.styles.Border(top=thin, left=thin, right=thin, bottom=thin)
        ws_sum.column_dimensions['Q'].width = 22.00
    
    for x in range(0, len(locs)):
        ws_sum['B'+str(x+26)] = locs[x]
        ws_sum['B'+str(x+26)].font = xl.styles.Font(name='Calibri', size=11, b=True)
        ws_sum['B'+str(x+26)].fill = xl.styles.PatternFill("solid", start_color='ffff00', end_color='ffff00')
        ws_sum['B'+str(x+26)].alignment = xl.styles.Alignment(horizontal="center", vertical="center")
    ws_sum.column_dimensions['B'].width = 7.00
    
    ws_sum['H42'].font = xl.styles.Font(name='Calibri', size=11)
    ws_sum['H42'].fill = xl.styles.PatternFill("solid", start_color='ffff00', end_color='ffff00')
    ws_sum['H42'].border = xl.styles.Border(top=thick, left=thick, right=thick, bottom=thick)
    ws_sum['H43'] = "Running Average"
    ws_sum['H43'].font = xl.styles.Font(name='Times New Roman', size=11, b=True)
    ws_sum['H43'].alignment = xl.styles.Alignment(horizontal="center", vertical="center")
    
    rows = [25,27,29,31,33,35,37,39]
    cols = ['D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q']
    for x in cols:
        for y in rows:
            ws_sum[x+str(y)].fill = xl.styles.PatternFill("solid", start_color='ddebf7', end_color='ddebf7')
    
    cell_range = ws_sum['D24':'Q40']
    for row in cell_range:
        for cell in row:
            cell.border = xl.styles.Border(top=thin, left=thin, right=thin, bottom=thin)
    
    cell_range = ws_sum['D2':'Q2']
    ws_sum.freeze_panes = 'D3'
    
    ## Summary Data frames
    ws_sum.merge_cells('T2:W2')
    ws_sum['T2'] = "Summary YTD"
    ws_sum['T2'].font = xl.styles.Font(name='Times New Roman', size=18, b=True)
    ws_sum['T2'].fill = xl.styles.PatternFill("solid", start_color='9bc2e6', end_color='9bc2e6')
    ws_sum['T2'].border = xl.styles.Border(top=thin, left=thin, right=thin, bottom=thin)
    ws_sum['U2'].border = xl.styles.Border(top=thin, left=thin, right=thin, bottom=thin)
    ws_sum['V2'].border = xl.styles.Border(top=thin, left=thin, right=thin, bottom=thin)
    ws_sum['W2'].border = xl.styles.Border(top=thin, left=thin, right=thin, bottom=thin)
    ws_sum['T2'].alignment = xl.styles.Alignment(horizontal="center", vertical="center")
    
    ws_sum['T3'] = cur_year
    ws_sum['T3'].font = xl.styles.Font(name='Times New Roman', size=11, color='FFFFFF', b=True)
    ws_sum['T3'].fill = xl.styles.PatternFill("solid", start_color='5b9ad5', end_color='5b9ad5')
    ws_sum['T3'].border = xl.styles.Border(top=thin, left=thin, right=thin, bottom=thin)
    ws_sum['T3'].alignment = xl.styles.Alignment(horizontal="center", vertical="center")
    ws_sum.column_dimensions['T'].width = 14.00
    ws_sum['U3'] = last_year
    ws_sum['U3'].font = xl.styles.Font(name='Times New Roman', size=11, color='FFFFFF', b=True)
    ws_sum['U3'].fill = xl.styles.PatternFill("solid", start_color='5b9ad5', end_color='5b9ad5')
    ws_sum['U3'].border = xl.styles.Border(top=thin, left=thin, right=thin, bottom=thin)
    ws_sum['U3'].alignment = xl.styles.Alignment(horizontal="center", vertical="center")
    ws_sum.column_dimensions['U'].width = 14.00
    ws_sum['V3'] = "Dollar Change"
    ws_sum['V3'].font = xl.styles.Font(name='Times New Roman', size=11, color='FFFFFF', b=True)
    ws_sum['V3'].fill = xl.styles.PatternFill("solid", start_color='5b9ad5', end_color='5b9ad5')
    ws_sum['V3'].border = xl.styles.Border(top=thin, left=thin, right=thin, bottom=thin)
    ws_sum['V3'].alignment = xl.styles.Alignment(horizontal="center", vertical="center")
    ws_sum.column_dimensions['V'].width = 16.00
    ws_sum['W3'] = "% Change"
    ws_sum['W3'].font = xl.styles.Font(name='Times New Roman', size=11, color='FFFFFF', b=True)
    ws_sum['W3'].fill = xl.styles.PatternFill("solid", start_color='5b9ad5', end_color='5b9ad5')
    ws_sum['W3'].border = xl.styles.Border(top=thin, left=thin, right=thin, bottom=thin)
    ws_sum['W3'].alignment = xl.styles.Alignment(horizontal="center", vertical="center")
    ws_sum.column_dimensions['W'].width = 14.00
    
    ws_sum.merge_cells('Y2:AB2')
    months = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]
    ws_sum['Y2'] = months[cur_latest_mon-1]
    ws_sum['Y2'].font = xl.styles.Font(name='Times New Roman', size=18, b=True)
    ws_sum['Y2'].fill = xl.styles.PatternFill("solid", start_color='9bc2e6', end_color='9bc2e6')
    ws_sum['Y2'].border = xl.styles.Border(top=thin, left=thin, right=thin, bottom=thin)
    ws_sum['Z2'].border = xl.styles.Border(top=thin, left=thin, right=thin, bottom=thin)
    ws_sum['AA2'].border = xl.styles.Border(top=thin, left=thin, right=thin, bottom=thin)
    ws_sum['AB2'].border = xl.styles.Border(top=thin, left=thin, right=thin, bottom=thin)
    ws_sum['Y2'].alignment = xl.styles.Alignment(horizontal="center", vertical="center")
    
    ws_sum['Y3'] = cur_year
    ws_sum['Y3'].font = xl.styles.Font(name='Times New Roman', size=11, color='FFFFFF', b=True)
    ws_sum['Y3'].fill = xl.styles.PatternFill("solid", start_color='5b9ad5', end_color='5b9ad5')
    ws_sum['Y3'].border = xl.styles.Border(top=thin, left=thin, right=thin, bottom=thin)
    ws_sum['Y3'].alignment = xl.styles.Alignment(horizontal="center", vertical="center")
    ws_sum.column_dimensions['Y'].width = 14.00
    ws_sum['Z3'] = last_year
    ws_sum['Z3'].font = xl.styles.Font(name='Times New Roman', size=11, color='FFFFFF', b=True)
    ws_sum['Z3'].fill = xl.styles.PatternFill("solid", start_color='5b9ad5', end_color='5b9ad5')
    ws_sum['Z3'].border = xl.styles.Border(top=thin, left=thin, right=thin, bottom=thin)
    ws_sum['Z3'].alignment = xl.styles.Alignment(horizontal="center", vertical="center")
    ws_sum.column_dimensions['Z'].width = 14.00
    ws_sum['AA3'] = "Dollar Change"
    ws_sum['AA3'].font = xl.styles.Font(name='Times New Roman', size=11, color='FFFFFF', b=True)
    ws_sum['AA3'].fill = xl.styles.PatternFill("solid", start_color='5b9ad5', end_color='5b9ad5')
    ws_sum['AA3'].border = xl.styles.Border(top=thin, left=thin, right=thin, bottom=thin)
    ws_sum['AA3'].alignment = xl.styles.Alignment(horizontal="center", vertical="center")
    ws_sum.column_dimensions['AA'].width = 16.00
    ws_sum['AB3'] = "% Change"
    ws_sum['AB3'].font = xl.styles.Font(name='Times New Roman', size=11, color='FFFFFF', b=True)
    ws_sum['AB3'].fill = xl.styles.PatternFill("solid", start_color='5b9ad5', end_color='5b9ad5')
    ws_sum['AB3'].border = xl.styles.Border(top=thin, left=thin, right=thin, bottom=thin)
    ws_sum['AB3'].alignment = xl.styles.Alignment(horizontal="center", vertical="center")
    ws_sum.column_dimensions['AB'].width = 14.00
    
    cell_range = ws_sum['T4':'W15']
    for row in cell_range:
        for cell in row:
            cell.border = xl.styles.Border(top=thin, left=thin, right=thin, bottom=thin)
    
    cell_range = ws_sum['Y4':'AB15']
    for row in cell_range:
        for cell in row:
            cell.border = xl.styles.Border(top=thin, left=thin, right=thin, bottom=thin)
    
    rows = [4,6,8,10,12,14]
    cols = ['T', 'U', 'V', 'W', 'Y', 'Z', 'AA', 'AB']
    for x in cols:
        for y in rows:
            ws_sum[x+str(y)].fill = xl.styles.PatternFill("solid", start_color='ddebf7', end_color='ddebf7')
    
    ## Enter the information into the right cells
    # Start with current year data frame
    # Work location by location, by month
    # Summary row for each location: current year table starts at row 4
    loc_nums_dict = {name: 4+i for i, (name, code) in enumerate(config_locations)}
    month_cols_dict = {"Jan": "D",
                       "Feb": "E",
                       "Mar": "F",
                       "Apr": "G",
                       "May": "H",
                       "Jun": "I",
                       "Jul": "J",
                       "Aug": "K",
                       "Sep": "L",
                       "Oct": "M",
                       "Nov": "N",
                       "Dec": "O"
                      }
    cur_sh_year = str(cur_year)[2:]
    ref_sheets = [x for x in wb.sheetnames if cur_sh_year in x]
    # Values for each location
    for x in range(0, len(locations)):
        #print("Loc: "+locations[x])
        ref_sheet = ref_sheets[x]
        row = loc_nums_dict[locations[x]]
        #print("Ref Sheet: "+ref_sheet)
        # Up to the current month
        for month in months[:cur_latest_mon]:
            col = month_cols_dict[month[:3]]
            #print("Col: "+str(col))
            val = find_month_total(wb[ref_sheet], month[:3])
            #print("Value: "+str(val))
            ws_sum[col+str(row)] = val
            ws_sum[col+str(row)].number_format = '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)'
        # Filling in the empty cells
        for month in months[cur_latest_mon:]:
            col = month_cols_dict[month[:3]]
            ws_sum[col+str(row)] = "=A1"
            ws_sum[col+str(row)].number_format = '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)'
        # Average/Location values
        ws_sum['P'+str(row)] = "=AVERAGE(D"+str(row)+":"+str(month_cols_dict[months[cur_latest_mon-1][:3]])+str(row)+")"
        ws_sum['P'+str(row)].number_format = '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)'
        # Total values for each location
        ws_sum['Q'+str(row)] = "=SUM(D"+str(row)+":"+str(month_cols_dict[months[cur_latest_mon-1][:3]])+str(row)+")"
        ws_sum['Q'+str(row)].number_format = '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)'
        # Summary data frame values
        ws_sum['T'+str(row)] = "=P"+str(row)
        ws_sum['T'+str(row)].number_format = '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)'
        # Month-to-date data frame values
        ws_sum['Y'+str(row)] = "="+month_cols_dict[months[cur_latest_mon-1][:3]]+str(row)
        ws_sum['Y'+str(row)].number_format = '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)'
    # Totals for each month
    for month in months[:cur_latest_mon]:
        col = month_cols_dict[month[:3]]
        ws_sum[col+"18"] = "=SUM("+col+"4:"+col+"17)"
        ws_sum[col+"18"].number_format = '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)'
    for month in months[cur_latest_mon:]:
        col = month_cols_dict[month[:3]]
        ws_sum[col+"18"] = "=A1"
        ws_sum[col+"18"].number_format = '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)'
    ws_sum['P18'] = "=SUM(P4:P17)"
    ws_sum['P18'].number_format = '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)'
    ws_sum['Q18'] = "=SUM(Q4:Q17)"
    ws_sum['Q18'].number_format = '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)'
    ws_sum['H20'] = "=AVERAGE(D18:"+month_cols_dict[months[cur_latest_mon-1][:3]]+"18)"
    ws_sum['H20'].number_format = '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)'
    
    # Move on to last year's data frame
    # Work location by location, by month
    # Summary row for each location: last year table starts at row 26
    loc_nums_dict = {name: 26+i for i, (name, code) in enumerate(config_locations)}
    month_cols_dict = {"Jan": "D",
                       "Feb": "E",
                       "Mar": "F",
                       "Apr": "G",
                       "May": "H",
                       "Jun": "I",
                       "Jul": "J",
                       "Aug": "K",
                       "Sep": "L",
                       "Oct": "M",
                       "Nov": "N",
                       "Dec": "O"
                      }
    last_sh_year = str(last_year)[2:]
    ref_sheets = [x for x in wb.sheetnames if last_sh_year in x]
    # Values for each location
    for x in range(0, len(locations)):
        #print("Loc: "+locations[x])
        ref_sheet = ref_sheets[x]
        row = loc_nums_dict[locations[x]]
        #print("Ref Sheet: "+ref_sheet)
        # Up to the current month
        for month in months[:last_latest_mon]:
            col = month_cols_dict[month[:3]]
            #print("Col: "+str(col))
            val = find_month_total(wb[ref_sheet], month[:3])
            #print("Value: "+str(val))
            ws_sum[col+str(row)] = val
            ws_sum[col+str(row)].number_format = '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)'
        # Filling in the empty cells
        for month in months[last_latest_mon:]:
            col = month_cols_dict[month[:3]]
            ws_sum[col+str(row)] = "=A1"
            ws_sum[col+str(row)].number_format = '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)'
        # Average/Location values
        ws_sum['P'+str(row)] = "=AVERAGE(D"+str(row)+":"+str(month_cols_dict[months[last_latest_mon-1][:3]])+str(row)+")"
        ws_sum['P'+str(row)].number_format = '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)'
        # Total values for each location
        ws_sum['Q'+str(row)] = "=SUM(D"+str(row)+":"+str(month_cols_dict[months[last_latest_mon-1][:3]])+str(row)+")"
        ws_sum['Q'+str(row)].number_format = '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)'
        # Summary data frame values
        ws_sum['U'+str(row-22)] = "=P"+str(row)
        ws_sum['U'+str(row-22)].number_format = '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)'
        # Month-to-date data frame values
        ws_sum['Z'+str(row-22)] = "="+month_cols_dict[months[cur_latest_mon-1][:3]]+str(row)
        ws_sum['Z'+str(row-22)].number_format = '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)'
    # Totals for each month
    for month in months[:last_latest_mon]:
        col = month_cols_dict[month[:3]]
        ws_sum[col+"40"] = "=SUM("+col+"26:"+col+"39)"
        ws_sum[col+"40"].number_format = '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)'
    for month in months[last_latest_mon:]:
        col = month_cols_dict[month[:3]]
        ws_sum[col+"40"] = "=A1"
        ws_sum[col+"40"].number_format = '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)'
    ws_sum['P40'] = "=SUM(P26:P37)"
    ws_sum['P40'].number_format = '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)'
    ws_sum['Q40'] = "=SUM(Q26:Q37)"
    ws_sum['Q40'].number_format = '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)'
    ws_sum['H42'] = "=AVERAGE(D40:"+month_cols_dict[months[last_latest_mon-1][:3]]+"40)"
    ws_sum['H42'].number_format = '_("$"* #,##0.00_);_("$"* \\(#,##0.00\\);_("$"* "-"??_);_(@_)'
    
    # Filling in the Summary data frames
    for x in range(4, 16):
        # Dollar Change
        ws_sum['V'+str(x)] = "=T"+str(x)+"-U"+str(x)
        ws_sum['V'+str(x)].number_format = '"$"#,##0_);[Red]\\("$"#,##0\\)'
        ws_sum['V'+str(x)].alignment = xl.styles.Alignment(horizontal="center", vertical="center")
        # % Change
        ws_sum['W'+str(x)] = "=IFERROR((T"+str(x)+"-U"+str(x)+")/U"+str(x)+",0)"
        ws_sum['W'+str(x)].number_format = '0%;[Red]\\(0%\\)'
        ws_sum['W'+str(x)].alignment = xl.styles.Alignment(horizontal="center", vertical="center")
        # Dollar Change
        ws_sum['AA'+str(x)] = "=Y"+str(x)+"-Z"+str(x)
        ws_sum['AA'+str(x)].number_format = '"$"#,##0_);[Red]\\("$"#,##0\\)'
        ws_sum['AA'+str(x)].alignment = xl.styles.Alignment(horizontal="center", vertical="center")
        # % Change
        ws_sum['AB'+str(x)] = "=AA"+str(x)+"/Z"+str(x)
        ws_sum['AB'+str(x)].number_format = '0%;[Red]\\(0%\\)'
        ws_sum['AB'+str(x)].alignment = xl.styles.Alignment(horizontal="center", vertical="center")
    
    # Open the spreadsheet to summary and let the user do with it as they will
    cwd = os.getcwd()
    wb.save(os.path.join(cwd, "Expense_Report.xlsx"))

    return 0

if __name__ == "__main__":
    main()