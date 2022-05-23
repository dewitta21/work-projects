# -*- coding: utf-8 -*-
"""
This script imports multiple .xlsx files containing object statistics exported from Avia

@date 07/15/2021
@ author Dr. Caleb Stoltzfus
"""
from tkinter import Tk
from tkinter.filedialog import askopenfilename
import pandas as pd
import os

# # Use tkinter to ask for the file path with a user interface
# Tk().withdraw() #supresses full GUI
# fnm = askopenfilename()

# indir = 'C://Users/caleb/My Drive/Lightspeed/Projects -active/202112 Pfizer/Objects/'
# indsave = 'C://Users/caleb/My Drive/Lightspeed/Projects -active/202112 Pfizer/Objects/cytomap/'

# indir = 'C://Users/caleb/My Drive/Lightspeed/Projects -active/202112 Pfizer/Objects/D18-2-1/'
# indsave = 'C://Users/caleb/My Drive/Lightspeed/Projects -active/202112 Pfizer/Objects/D18-2-1/'

indir = 'C://Users/caleb/My Drive/Lightspeed/Projects-active/202112 Pfizer/Objects/D25-11-1/'
indsave = 'C://Users/caleb/My Drive/Lightspeed/Projects-active/202112 Pfizer/Objects/D25-11-1/'

savetype  = 'merge'
# savetype  = 'single'

# input the size of a pixel in micrometers
pxdim = 0.441

# Load in the .xls file
# file_xls = pd.ExcelFile(fnm)

#  Get the number of sheets (channels) fromt he file
# NCh = len(file_xls.sheet_names)-1


# Import the measurments file and format it
infiles = os.listdir(indir)

f=0
for infile in infiles:
    if not infile.lower().endswith('.xlsx'):
        continue
    print( infile)
    # Load in the .xls file
    file_xls = pd.ExcelFile(indir + infile)
    fnmxlsx = 'Merged_' + infile
    infile
    # Build a dataframe
    n=0
    for shtnm in file_xls.sheet_names:
        if shtnm != 'Summary' and shtnm != 'Surfaces.Summary' and not shtnm.startswith('Cross Sections.'): #end of condition
            df = pd.read_excel(file_xls, shtnm) #pull the sheet
            #Define the column name
            colmnm = df.columns[0]
            # Take out some common name errors
            colmnm = colmnm.replace('Surfaces.', '')
            print(colmnm)
            if n == 0:                           
                dat = pd.DataFrame(data=df[df.columns[0]].values, columns={'Mesh'}) #add the data to the dataframe                                
                dat.Mesh=dat.Mesh.str.replace('Mesh', '')
                dat[colmnm] = df[df.columns[1]] 
                n=n+1
            else :   
                dat[colmnm] = df[df.columns[1]] 
    
    print(' ')
    
    if 'Crop_' in infile :  
        # find the index of the crop position
        idxcrp = infile.find('Crop_')+5
        idxcrpend = infile[idxcrp::].find('-')
        crpoffset = infile[idxcrp:(idxcrp+idxcrpend)]
        print('Cropped offset = ' + crpoffset)
        crpoffset = int(crpoffset)
    
    # scale and shift any centriod measurments 
    if 'Centroid X' in list(dat.columns.values) :
        # If this data is from aivia 8 or earlier position is in pixels
        
        # check to see if this is a cropped image
        if 'Crop_' in infile :  
            print('Z offset by ' + str(crpoffset))
            dat['Centroid Z'] = dat['Centroid Z'] + crpoffset
            
        # scale aivia 8 data from pixels to micrometers
        dat['Centroid X'] = dat['Centroid X']*(pxdim)
        dat['Centroid Y'] = dat['Centroid Y']*(pxdim)
        dat['Centroid Z'] = dat['Centroid Z']*(pxdim)
        
    else :
        # check to see if this is a cropped image
        if 'Crop_' in infile :  
            print('Z offset by ' + str(crpoffset))
            dat['Center of Mass Z (µm)'] = dat['Center of Mass Z (µm)']+(crpoffset*pxdim)  
        
    # rename some columns  and swap x and z  
    dat=dat.rename(columns={'Mean Intensity - CXCR9_647' : 'CXCL9', 'Mean Intensity - CXCL3_561' : 'CXCR3'})
    dat=dat.rename(columns={'Centroid X' : 'Position_Z', 'Centroid Y': 'Position_Y', 'Centroid Z': 'Position_X'}) 
    dat=dat.rename(columns={'Center of Mass X (µm)' : 'Position_Z', 'Center of Mass Y (µm)': 'Position_Y', 'Center of Mass Z (µm)': 'Position_X'})
    dat=dat.rename(columns={'Mean Intensity - CXCL9_647' : 'CXCL9', 'Mean Intensity - CXCR3_561': 'CXCR3', 'Mean Intensity - CD8_488': 'CD8'})
    dat=dat.rename(columns={'Average Intensity - CXCL9_647' : 'CXCL9', 'Average Intensity - CXCR3_561': 'CXCR3', 'Average Intensity - CD8_488': 'CD8'})
    dat=dat.rename(columns={'Mean Intensity - CD8_cells' : 'CD8_cells'})
    dat=dat.rename(columns={'Volume (µm³)' : 'Volume'})
    
    if 'CD8_cells' in list(dat.columns.values) :
        column_names = ["Position_X", "Position_Y", "Position_Z", 'Sphericity', 'Volume', 'Mesh', 'CXCL9', 'CXCR3', 'CD8', 'CD8_cells']
    elif 'Center Line Length (µm)' in list(dat.columns.values) :
        column_names = ["Position_X", "Position_Y", "Position_Z", 'Sphericity', 'Volume', 'Mesh', 'CXCL9', 'CXCR3', 'CD8', 'Surface Area To Volume Ratio', 'Volume Ratio', 'Center Line Length (µm)'] 
    else :
        column_names = ["Position_X", "Position_Y", "Position_Z", 'Sphericity', 'Volume', 'Mesh', 'CXCL9', 'CXCR3', 'CD8']
    if savetype  == 'merge' :
        column_names = ["Position_X", "Position_Y", "Position_Z", 'Sphericity', 'Volume', 'Mesh', 'CXCL9', 'CXCR3', 'CD8']
 
        
    dat = dat.reindex(columns=column_names)
    
    # add cell type classifier columns
    if 'CD8' in infile :  
        dat['Classifier'] = dat['Position_X']*0+1
    elif 'CXCL9' in infile :
        dat['Classifier'] = dat['Position_X']*0+2
    elif 'CXCR3' in infile :
        dat['Classifier'] = dat['Position_X']*0+3
    else :
        dat['Classifier'] = dat['Position_X']*0

    print(list(dat.columns.values))
    
##############################                
    # If you want to convert a bunch of files use this
    # if it is the first file build the table
    # Now I have a table I can save it
    
    # fnm_write = os.path.splitext(os.path.basename(infile))[0] +'.csv'
    # fnm_write = fnm_write.replace('-', '_')
    # fnm_write = indsave + fnm_write
    # dat.to_csv(fnm_write, index=False)
    # print(fnm_write)
    # print(' ')
##############################


##############################
# If you want to concatenate a bunch of files use this
    if f==0:
        # if it is the first file build the table
        # Now I have a table I can save it
        fnm_write = '/' + os.path.splitext(os.path.basename(fnmxlsx))[0] +'.csv'
        fnm_write = fnm_write.replace('-', '_')
        dat.to_csv(indsave + fnm_write, index=False, float_format='%.5f')
        print(indsave + fnm_write)
        f=f+1
    else :
        # if it is any other file append to the table
        dat.to_csv(indsave + fnm_write , index=False , mode='a', header=False, float_format='%.5f')
############################## 
    print(' ')

                      

                      
