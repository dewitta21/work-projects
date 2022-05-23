# -*- coding: utf-8 -*-
"""
This should open an image, and show the volume of an image as a total and as a funciton of Z
The idea is to use the pixel classifier to define some feature (like collagen) 
then use this approach to do a simple volume measurement 


Created on 9/16/2021
@author: Dr. Caleb Stoltzfus
"""
# Make sure parent folder is in working directory
import sys
import os
# import subprocess

wd = os.path.abspath(os.path.join(os.path.dirname(__file__), os.pardir))
syspth = sys.path
if wd not in syspth:
    sys.path.append(wd)

# # import the necesary external libraries
# from tkinter import Tk
# from tkinter.filedialog import askopenfilename


## Define the internal functions


from plyfile import PlyData
import pandas as pd

########################
########################

# # Use tkinter to ask for the file path with a user interface
# Tk().withdraw() #supresses full GUI
# fnm = askopenfilename()

indir = 'C://Hirschprung/excel/SCH4-A/'
indsave = 'C://Hirschprung/other/'

# define the number of micrometers per pixel
dim = 0.441

# Import the measurments file and format it
infiles = os.listdir(indir)
for infile in infiles:
    if not infile.lower().endswith('.xlsx'):
        continue
    print( infile)
    # Load in the .xls file
    file_xls = pd.ExcelFile(indir + infile)
    fnmxlsx = infile
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
                dat[colmnm] = df[df.columns[1]] 
                n=n+1
            else :   
                dat[colmnm] = df[df.columns[1]] 

infiles = os.listdir(indir)

# infiles = ['Mesh 2.ply', 'Mesh 5.ply', 'Mesh 193.ply', 'Mesh 1000.ply', 'Mesh 1280.ply']
dat=dat.rename(columns={'Mean Intensity - PGP9.5_561' : 'PGP9.5'})
n=0
for infile in infiles:
    if not infile.lower().endswith('.ply'):
        # print( "skipping %s" % infile)
        continue
    # print( infile)
    plydata = PlyData.read(indir + infile)
    
    # vertices
    if plydata.elements[0].name=='vertex':
        
        # load the vertices       
        vertices = pd.DataFrame(data=plydata.elements[0].data, columns={'y', 'z', 'x'})
        
        # rename the positional columns so they are gmore generally recognizable
        vertices = vertices.rename(columns={'x':'Position_X', 'y':'Position_Y', 'z':'Position_Z'})
        
        # convert to micrometers
        vertices = vertices*dim

        # add a column with mesh name
        vertices['Mesh'] = infile[:-4]
        
        # Downsample vertices
        if max(vertices.shape) > 100 :
            # take every fourth vertex
            vertices = vertices[0::50]
        
        
            
        if n==0 :
    
            # merge the vertices with the calculations
            datFULL = vertices.merge(dat, on=['Mesh'])
            
            # Now I have a table I can save it
            fnm_write = '/' + os.path.splitext(os.path.basename(fnmxlsx))[0] +'.csv'
            fnm_write = fnm_write.replace('-', '_')
            datFULL.to_csv(indsave + fnm_write , index=False, float_format='%.5f')

            n=1
        else :

            # merge the vertices with the calculations
            datsub = vertices.merge(dat, on=['Mesh'])
            # datFULL = datFULL.append(datsub)
            
            # Now I have a table I can save it
            datsub.to_csv(indsave + fnm_write , index=False , mode='a', header=False, float_format='%.5f')
            
    else :
        print('formatting error')
            
    print( infile)
    
    
# # Now I have a table I can save it
# fnm_write = '/' + os.path.splitext(os.path.basename(fnmxlsx))[0] +'.csv'
# # path_write = os.path.join(os.path.dirname(fnm), fnm_write)
# datFULL.to_csv(indsave + fnm_write , index=False)
            
