#!/usr/bin/env python3

import pandas as pd
import numpy as n
import openpyxl as xl
import numpy as np
#create well labels

rows = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H']
cols = np.arange(1, 13, 1).tolist()


def ctlimport(
        workbooks,
        ctl,
        export):

    # load all CTL .xlsx files, generate plateID strings from their filenames.
    ctls = [xl.load_workbook(excel).active for excel in ctl]

    stuff = [sheet.iter_cols(min_col=3, max_col=14, min_row=55, max_row=62) for sheet in ctls] 

    # load worksheet containing sample names. 
    samplesheets = [xl.load_workbook(workbook)['Serum Dilution'] for workbook in workbooks]
    # dates = [str(workbook).partition('/')[2] for workbook in workbooks
    #          if '/' in workbook]
    # dates = [str(date).partition('.')[0] for date in dates]

    numofplates = len(ctls) 

    labels = []
    for i in np.arange(0, numofplates, 1):
        for col in cols:
            for row in rows:
                labels.append(row + str(col))

    #samplexlsx = xl.load_workbook('2023-May-23 FFA RSVNeut.xlsx')['Serum Dilution'] 

    # plateIDs = []
    # for platenum in np.arange(1, numofplates+1, 1):
    #     for i in np.arange(0, 96, 1):
    #         plateIDs.append(f'plate {platenum} {plateID}')

    plateIDs = [str(countsheet).replace('.xlsx', '') for countsheet in ctl]

    num_wells = np.arange(2, (numofplates * 96) + 1, 1).tolist()

    #create new workbook, sheet, and label columns
    port = xl.Workbook()
    new = port.active

    new['A1'].value = 'foci_num'
    new['B1'].value = 'fold_dil'
    new['C1'].value = 'type'
    new['D1'].value = 'wellID'
    new['E1'].value = 'sample_num'
    new['F1'].value = 'plateID'

    # BEGIN EXPORT: read only rows in neut worksheet that have values. 
    
    print(numofplates)
    for samplexlsx in samplesheets:
        if numofplates == 7:
            sumofsamples = [sample for sample in samplexlsx['C'][1:samplexlsx.max_row]
                            if sample.value != None]
        else:
            sumofsamples = [sample for sample in samplexlsx['C'][1:4*numofplates+1]]

    # for i, sample in enumerate(sumofsamples):
    #     if '/' in sample.value:
    #         sumofsamples[i].value = str(sample.value).replace('/', '.')
    #     else:
    #         pass

    samplelist = []
    for i in np.arange(0, 2, 1):
        for sample in sumofsamples:
            samplelist.append(sample)
            samplelist.append(sample)

    #twice the samples, but each will fill only 1 column (half the coverage)
    #one chunk = 1 plate

    chunk_size = 8
    chunked_list = [samplelist[i:i+chunk_size] for i in range(0, len(samplelist), chunk_size)]

    # create dilution labels
    dils = []
    for i in np.arange(0, 2, 1):
        for group in chunked_list:
            for sample in group: 
                dils.append(20)
                dils.append(60)
                dils.append(180)
                dils.append(540)
                dils.append(1620)
                dils.append(4860)

    sampleids = []
    for index, i in  enumerate(np.arange(0, len(ctl), 1)):
        for sampid in np.arange(0, 5, 1):
            sampleids.append(i)
            sampleids.append(i)
            sampleids.append(i)
            sampleids.append(i)
            sampleids.append(i)
            sampleids.append(i)
    
    dils = [dils[i:i+6] for i in range(0, len(dils), 6)]

    # negatives occupy rows A and H, and columns 1 and 12.
    negs = [num for num in np.arange(2, 96*numofplates + 1, 1).tolist()
            if num > 9 if (num-1)% 8 == 0 or (num-2)% 8 == 0]

    for a in np.arange(2, 10, 1):
        for b in np.arange(0, numofplates, 1):
            negs.append((a + 96*b))

    for a in np.arange(89, 98, 1):
        for b in np.arange(0, numofplates, 1):
            negs.append((a+96*b))

    negs.sort()


    # VOCs occupy rows B-G of columns 10 and 11. 
    vocs = []

    for a in np.arange(74,80,1):
        for b in np.arange(0, numofplates, 1):
            vocs.append((a+96*b)+1)

    for a in np.arange(82,88, 1):
        for b in np.arange(0,numofplates,1):
            vocs.append((a+96*b)+1)

    vocs.sort()

    # create indices for wells that *aren't* negatives or VOCs
    samplenums = [num for num in num_wells
        if (num not in negs) and (num not in vocs)]

    both_sample_cols = [samplenums[i:i+12] for i in range(0, len(samplenums), 12)]

    platebyplate = [both_sample_cols[i:i+4] for i in range(0, len(both_sample_cols), 4)]

    each_sample_col = [samplenums[i:i+6] for i in range(0, len(samplenums), 6)]


    # fill new sheet with counts from CTL file
    for index, sheet in enumerate(stuff):
        for i, oldcol in enumerate(sheet, 0):
            for row, count in enumerate(oldcol):
                new.cell(row=(row+(8*i))+2+(96*index), column=1).value = count.value

    # fill new sheet with sample names from neut assay worksheet
    for cellnums, sampleid in zip(each_sample_col, samplelist):
        for cell in new['C'][cellnums[0]-1:cellnums[5]]:
            cell.value = sampleid.value

    for dil, col in zip(dils, each_sample_col):
        for d, cell in zip(dil, new['B'][col[0]-1:col[5]]):
            cell.value = d 
    
    for i, plate in enumerate(platebyplate):
        for both_cols in platebyplate:
            i = 0
            for col in both_cols:
                i += 1
                for cell in new['E'][col[0]-1:col[5]]:
                    cell.value = i
                for cell in new['E'][col[6]-1:col[11]]:
                    cell.value = i
                    

    # fill new sheet with well labels
    for label, row in zip(labels, new.iter_rows(min_row=2, max_row=96*numofplates+1, min_col =4,  max_col=4)):
        for cell in row:
            cell.value = label

    # fill new sheet with plate IDs
    for i, plateID in enumerate(plateIDs, 1):
        # for cell in new['F'][i+(96*(i-1)):i*96+1]:
        for cell in new['F'][1+(96*(i-1)): 97*(i)]:
            cell.value = plateIDs[i-1]

# the sample in every 1st and 8th row of every 8 rows is a negative. Annotate it as such.
    for i in np.arange(0, len(negs), 1):
        new.cell(row=negs[i], column=3).value = 'negative'

# annotate VOC
    for i in np.arange(0, len(vocs), 1):
        new.cell(row=vocs[i], column=3).value = 'VOC'


    for type, dil in zip(new['C'][0:96*numofplates + 1], new['B'][0:96*numofplates + 1]):
        if type.value == 'VOC':
            dil.value = ''
        else:
            pass


    for type, samplenum in zip(new['C'][0:96*numofplates + 1], new['E'][0:96*numofplates + 1]):
        if type.value == 'VOC':
            samplenum.value = ''
        else: 
            pass

    port.save(str(export)+'.xlsx')





