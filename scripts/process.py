#!/usr/bin/python
# -*- coding: utf8 -*-

import os
import urllib
import csv
import xlrd


source = 'http://sedac.ipcc-data.org/ddc/observed/data/IPCC_AR4_Observed_CC_Impacts_Database_v1.0.xls'
descriptions = {
    'Ch_Reg': {
        1: 'Africa',
        2: 'Asia',
        3: 'Australia and New Zealand',
        4: 'Europe',
        5: 'Latin America',
        6: 'North America',
        7: 'Polar Regions (Arctic and Antarctic)',
        8: 'Small Islands and Oceans '
    },
    'Code': {
        1: 'cryosphere',
        2: 'hydrology',
        3: 'coastal processes',
        4: 'marine and freshwater biological systems',
        5: 'terrestrial biological systems',
        6: 'agriculture and forestry',
        7: 'human health (no obs)',
        8: 'disasters and hazards (no obs)',
        9: 'socio-economic indicators (no obs)'
    },
    'SGNFCNT': {
        0: 'no, the impact observed is not statistically significant',
        1: 'yes, the impact observed is statistically significant',
        2: 'statistics not applied'
    },
    'Change': {
        1: 'change consistent with warming (with respect with HadCRUT3  temperatures)',
        2: 'change consistent with cooling',
        0: 'observation with no change'
    },
    'Air_T': {
        0: 'no, not statistically related',
        1: 'yes, statistically related',
        2: 'relationship not mentioned'
    },
    'SST': {
        0: 'no, not statistically related',
        1: 'yes, statistically related',
        2: 'relationship not mentioned'
    },
    'Prec': {
        0: 'no',
        1: 'yes',
        2: 'relationship not mentioned'
    },
    'AO': {
        0: 'no',
        1: 'yes',
        2: 'relationship not mentioned'
    },
    'NAO': {
        0: 'no, not statistically related',
        1: 'yes, statistically related',
        2: 'relationship not mentioned'
    },
    'ENSO': {
        0: 'Authors discount El Nino',
        1: 'Authors say El Nino is a factor',
        2: 'relationship not mentioned'
    },
    'Land': {
         0: 'Authors Discounted Land Use Change as a factor',
         1: 'Authors suggest land use was a contributing factor',
         2: 'relationship not mentioned'
    }
}

def setup():
    '''Crates the directorie for archive if they don't exist
    
    '''
    if not os.path.exists('archive'):
        os.mkdir('archive')

def retrieve(source):
    '''Downloades xls data to archive directory
    
    '''
    urllib.urlretrieve(source,'archive/external-data.xls')

def get_data():
    '''Gets the data from xls file and yields lists of it's data row by row
    
    '''
    countries = {}
    fo = xlrd.open_workbook('archive/external-data.xls')
    sheet = fo.sheet_by_index(2) 
    num_rows = sheet.nrows
    num_cols = sheet.ncols
    for i in range(1, num_rows):
        row = {}
        for n in range(1, num_cols):
            if n != 3: #4th colmn has unicode errors in several rows
                col_name = sheet.cell_value(0,n)
                value = sheet.cell_value(i,n)
                if col_name in descriptions:
                    row[col_name] = descriptions[col_name][value]
                else:
                    row[col_name] = value
        yield row

def process(data):
    '''takes generator funtion as input and writes data into csv file
    
    '''
    fo = open('data/climate-change.csv', 'w')
    fieldnames = [u'Year_Pub',
              u'Ref',
              u'Ch_Reg',
              u'Lat',
              u'Long',
              u'Yr_start',
              u'Yr_end',
              u'Duration',
              u'Category',
              u'Code',
              u'SGNFCNT',
              u'Change',
              u'Air_T',
              u'SST',
              u'Prec',
              u'AO',
              u'NAO',
              u'ENSO',
              u'Land']
    writer = csv.DictWriter(fo, fieldnames=fieldnames)
    writer.writeheader()
    for row in data:
        writer.writerow(row)
    fo.close()
        
if __name__ == '__main__':
    setup()
    retrieve(source)
    process(get_data())