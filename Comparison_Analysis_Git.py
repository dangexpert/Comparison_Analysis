# Comparison Analysis between two files. Implemented with graphical user interface (GUI). 
# Documents changes based on changed values, added values, or dropped values. 

import pandas as pd
import PySimpleGUI as sg
import openpyxl

# Added Graphical Interface -- Click Browse  to get Old & New File
layout = [[sg.Text('Select File to Compare: ')],    
                 [sg.Text('File 1', size=(8, 1)), sg.InputText("Insert Old File"), sg.FileBrowse()],      
                 [sg.Text('File 2', size=(8, 1)), sg.InputText("Insert New File"), sg.FileBrowse()], 
                 [sg.Text('Difference', size=(8, 1)), sg.InputText()],
                 [sg.Submit(), sg.Cancel()]]      

window = sg.Window('Comparison Analysis').Layout(layout)

event, values = window.Read() 
window.Close()
print(event, values) 

# Defines the difference between the two files in the created spreadsheet
def filedifference(x):
    return x[0] if x[0] == x[1] else '{} ---> {}'.format(*x) 
# Will add highlight option to easily see difference between cells 

# Refers to input in GUI to get file source to compare new & old
old = pd.read_excel(values[0], 'External', na_values=['NA'])
new = pd.read_excel(values[1], 'External', na_values=['NA'])

old['version'] = "old"
new['version'] = "new"

old.head()
new.head() 

# Uses a fixed key to compare the files. In this case it would be the loannumber
old_ln_all = set(old['LoanNumber'])
new_ln_all = set(new['LoanNumber'])

# Analyze fields to see if there is any dropped or added fields 
dropped = old_ln_all - new_ln_all
added = new_ln_all - old_ln_all

#join all the data together and ignore indexes so it all gets concatenated 
all_data = pd.concat([old,new],ignore_index=True)
all_data.head() 

# set main column names that matter within the datasets 
changes = all_data.drop_duplicates(subset=["Column1", " Column2" , "Column3"], keep='last') #insert your own column names
changes.head() 

#get all the duplicate rows 
dupe_accts = changes[changes['LoanNumber'].duplicated() == True]['LoanNumber'].tolist()
dupes = changes[changes["LoanNumber"].isin(dupe_accts)]
dupes 

# Pull out the old and new data into separate dataframes
change_new = dupes[(dupes["version"] == "new")]
change_old = dupes[(dupes["version"] == "old")]

# Drop the temp columns - we don't need them now
change_new = change_new.drop(['version'], axis=1)
change_old = change_old.drop(['version'], axis=1)

# Index on the loan numbers (Change to your own)
change_new.set_index('LoanNumber', inplace=True)
change_old.set_index('LoanNumber', inplace=True)

# Combine all the changes together
df_all_changes = pd.concat([change_old, change_new],
                            axis='columns',
                            keys=['old', 'new'],
                            join='outer')

df_all_changes = df_all_changes.swaplevel(axis='columns')[change_new.columns[0:]]
df_all_changes

df_changed = df_all_changes.groupby(level=0, axis=1).apply(lambda frame: frame.apply(filedifference, axis=1))
df_changed = df_changed.reset_index()
df_changed

# creates a list of removed and added items 
df_removed = changes[changes["LoanNumber"].isin(dropped)]
df_added = changes[changes["LoanNumber"].isin(added)]
df_added 

# save the changes to excel but only include the columns we care about 
output_columns = ["Column1", " Column2" , "Column3"]

# Creates new spreadsheet using your difference input name 
writer = pd.ExcelWriter("Insert Path Directory" + values[2] + ".xlsx") #use double slash or "R" infront of path

# Creates 3 sheets based on what's changed, removed, or added. 
df_changed.to_excel(writer,"Changed", index=False, columns=output_columns)
df_removed.to_excel(writer,"Removed",index=False, columns=output_columns)
df_added.to_excel(writer,"Added",index=False, columns=output_columns)
writer.save()

#formats spreadsheets width 
newFile = ("Insert Path Directory" + values[2] + ".xlsx") #use double slash or "R" infront of path

wb = openpyxl.load_workbook(filename=newFile)
worksheet = wb.active #activates the worksheet -- need to make sure to activate wb.save() at the end if changing excel 

for col in worksheet.columns:
    max_length = 0
    column = col[0].column

    for cell in col:
        try: 
#based on the value of the cell, it makes sure it equals the max_length 
            if len(str(cell.value)) > max_length:
                max_length = len(cell.value)
        except:
                pass
        adjusted_width = (max_length + 2) * 1.2
        worksheet.column_dimensions[column].width = adjusted_width

wb.save(newFile)