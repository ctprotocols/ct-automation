# -*- coding: utf-8 -*-
"""
Created on Thu Aug 13 07:03:00 2020

@author: EastmanE
"""
# =============================================================================
# If the program is stopped before line 100 or runs into an error before the copying
# is complete, run the command 'exc.Quit'. Then, open task manager and quit
# Excel. 
# =============================================================================
#####Import packages

def copy_function(blank_start):
    #####Import packages
    import os
    import win32com.client as win32
    from CTQC_settings_2021 import chartpath, chartnames, box_path, Z_path
    
    # =============================================================================
    # Set Dates
    # =============================================================================
    #Month name in 3 letter format (e.g. Aug)   
    month_shortname= blank_start.strftime('%b')    
    #Year in YYYY format            
    year = blank_start.strftime('%Y')                
    #YYYY-mm          
    year_month = blank_start.strftime('%Y-%m')        
    #YYYY-mm August        
    
    # =============================================================================
    # Dispatch Excel and create the destination file
    # =============================================================================
    #Destination folder name in Z drive
    destination_folder = year_month + ' CT QC'                     
    #Destination file name. This is what the program will name the file it creates.
    destination_file = 'CT QC ' + year_month + '.xlsx'          
    #Full path to the destination
    destination_path = os.path.join(Z_path, destination_folder, destination_file)    
    
    #Start Excel
    excel = win32.Dispatch('Excel.Application')              
    #Excel creates a new workbook
    CTQC_wb = excel.Workbooks.Add()       
    #Turns off any prompts from Excel (e.g. do you want to overwrite existing file?)
    excel.DisplayAlerts = False       
    #Saves the new workbook to the path. 
    CTQC_wb.SaveAs(destination_path)               
    
    # =============================================================================
    # Find paths
    # =============================================================================
    #Empty list to hold paths
    box_paths_list = []                                      
    #Look for directories in Box QA folder
    for QA_dir in os.listdir(box_path):                             
        #Add to the path
        path_L1 = os.path.join(box_path, QA_dir)                    
        #If it is a folder and not a file, List directories in each machine's folder
        if not os.path.isfile(path_L1):                             
            for QA_dir_L2 in os.listdir(path_L1):
                #Create the paths to each directory
                path_L2 = os.path.join(path_L1, QA_dir_L2)          
                #If the path leads to a file and the YYYY_MM is in the name, Add to the list (Most machines have this format)
                if (os.path.isfile(path_L2)) & (year_month in QA_dir_L2):   
                    box_paths_list.append(path_L2)                          
                #If the path leads to a folder and YYYY is in the folder name (Machine > 2020 folder > file)
                elif (not os.path.isfile(path_L2)) & (year in QA_dir_L2): 
                    #List directories in the YYYY folder
                    for QA_dir_L3 in os.listdir(path_L2):
                        #Find the YYYY_MM in the file name (MGB)
                        if year_month in QA_dir_L3:                         
                            #Create path, append path
                            path_L3 = os.path.join(path_L2, QA_dir_L3)      
                            box_paths_list.append(path_L3)                  
                #Find the AHSP file and append it
                elif (os.path.isfile(path_L2)) & (year in QA_dir_L2) & ('AHSP' in QA_dir_L2):   
                    box_paths_list.append(path_L2)
                else:
                    pass
        else:
            pass
    
    # =============================================================================
    # Copy QC worksheets
    # =============================================================================
    for path in box_paths_list:
        #open the worksheet
        machine_wb = excel.Workbooks.Open(Filename = path)                
        #AHSP worksheetname = 3 letters of the month all caps
        if 'AHSP' in path:
            machine_ws = machine_wb.Worksheets(month_shortname.upper())
        #All others are name "QA"
        else:
            machine_ws = machine_wb.Worksheets('QA')                                   
        #Copy the worksheet into the new workbook
        machine_ws.Copy(Before= CTQC_wb.Worksheets(1))             
        #Rename the sheet in the new workbook
        new_sheetname = path.split(os.sep)[6]
        CTQC_wb.Worksheets(1).Name = new_sheetname
        #Close the original worksheet and don't save any changes
        machine_wb.Close(SaveChanges=False)        
        #Print name to show progress
        print(new_sheetname)        
        
    # # =============================================================================
    # # Copy Linearity worksheets
    # # =============================================================================
    # #create a list of paths for Tosh/Canon machines (they have Lin tests as well)
    lin_paths = [path for i, path in enumerate(box_paths_list) if ('Toshiba' in path) | ('Canon' in path)]
    
    for i, lin_file_path in enumerate(lin_paths):
        #Open up the workbook
        lin_wb = excel.Workbooks.Open(Filename = lin_file_path)                      
        #Navigate to the lineariy worksheet, worksheet in workbook must be named "Toshiba Linearity"
        lin_ws = lin_wb.Worksheets('Toshiba Linearity')                       
        #Copy and rename
        lin_ws.Copy(Before= CTQC_wb.Worksheets(1))                               
        lin_name = lin_file_path.split(os.sep)[6]
        CTQC_wb.Worksheets(1).Name = lin_name + ' Linearity'                 
        #Close the workbook from box and don't save any changes to original workbook
        lin_wb.Close(SaveChanges=False)                                      
        
    # =============================================================================
    # Copy summary graphs from the template into the spreadsheet
    # =============================================================================
    #Open workbook containing templates for the three charts
    chartwb = excel.Workbooks.Open(Filename = chartpath) 
    
    #Copy the three charts. Chart names are defined in the settings file 
    for names in chartnames:                                
        chartws = chartwb.Worksheets(names)                 
        chartws.Copy(Before= CTQC_wb.Worksheets(1))    
        #Change the link for these worksheets to the current worksheet.         
        CTQC_wb.ChangeLink(Name = chartpath, NewName= destination_path)    
        #Rename the sheets
        CTQC_wb.Worksheets(1).Name = names                      
    
    # =============================================================================
    # Save/Close/Quit
    # =============================================================================
    #Delete Sheet1. This sheet is automatically created when you first create the file.     
    CTQC_wb.Worksheets('Sheet1').Delete()   
    CTQC_wb.Close(SaveChanges=True)         
    #Quit Excel. If the program gets stopped before this line, run this command to quit Excel.
    excel.Quit()       
                   
# import datetime as dt
# copy_function((dt.datetime.today()-dt.timedelta(days=30)).replace(hour = 0, minute = 0, second = 0, microsecond = 0))