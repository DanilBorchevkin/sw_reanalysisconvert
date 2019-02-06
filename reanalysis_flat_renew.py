# -*- coding: utf-8 -*-
"""
Created on Wed Aug  8 22:26:15 2018

@author: Danil Borchevkin

Dimension of the data 
1-42 by lat (rows)
1-50 by long (columns)

"""

from openpyxl import load_workbook
import csv


def drange_up(start, stop, step):
    ret = []
    r = start
    while r <= stop:
        ret.append(r)
        r += step
    return ret

def drange_down(start, stop, step):
    ret = []
    r = start
    while r >= stop:
        ret.append(r)
        r += step
    return ret

# Because reading of lats and longs very high load task we define it
LAT_LIST = drange_down(90, -90, -0.5)
LONG_LIST = drange_up(-180, 179.375, 0.625)
SHIFT = 2

def get_indicies_of_values(list_of_values, start_value, end_value, shift=0):
    indicies_list = []

    for i,value in enumerate(list_of_values):
        if start_value <= value <= end_value:
            indicies_list.append(i + shift)

    return indicies_list

def parse__reanalysis_file(path, filename, start_lat, end_lat, start_long, end_long):
    # TODO change workdir
    workbook = load_workbook(filename=path+filename, read_only=True)
    
    sheets = workbook.sheetnames
    
    for sheet in sheets:
        print('Work with sheet [' + sheet + ']')
        
        # 0. Activate sheet
        worksheet = workbook[sheet]

        # 1. Filter lats and longs according to params and save it's indexes inside sheet
        desired_lats_index = get_indicies_of_values(LAT_LIST, end_lat, start_lat)
        desired_longs_index = get_indicies_of_values(LONG_LIST, start_long, end_long)
        
        # 2. Create output list
        output_list = [] 
        for long_i in desired_longs_index:
            for lat_i in desired_lats_index:
                output_list.append([
                    LAT_LIST[lat_i],                                               #lat
                    LONG_LIST[long_i],                                             #long
                    worksheet.cell(row=lat_i+SHIFT, column=long_i+SHIFT).value     #value
                ])

        # 3. Write values to CSV
        with open('./' + filename + '_' + sheet + '.csv', 'w', newline='') as f:
            writer = csv.writer(f, delimiter='\t')
            for line in output_list:
                writer.writerow(line)

if __name__ == "__main__":
    print('Script is started')

    # 1. Change start and end values here
    # WARNING! Please see source file
    # - lats starts from 90 and ends by -90
    # - longs starts from -180 and ends by -179.375
    START_LAT = 90
    END_LAT = 85
    START_LONG = -180
    END_LONG = -170

    # 2. CHANGE FILENAME HERE. FILE MUST TO BE IN THE SAME FOLDER
    parse__reanalysis_file('./', 'test_renew.xlsx', START_LAT, END_LAT, START_LONG, END_LONG) 

    # 3. Enjoy your files in the same directory =))

    print('script is ended')