import pywildcard
import os
import shutil
import datetime
import time
import xlwings as xw

#backup original request overview file
#get current date to append to copy's name
today = str(datetime.datetime.now().date())
loc = r"C:\Users\shemchen\Desktop\excelPython"
overviewfile = "NS-OCI_Resource Management-v2"

max_rows = 500
four_nine_five_k = 495784
one_eight_four_k = 184374


#make a backup of the Total Request Overview excel sheet
def make_copy():
    shutil.copy(loc + "\\" + overviewfile + ".xlsx", loc + "\\" + overviewfile + " " + today + ".xlsx")

#find new files and insert them into a list
def find_new_files(location, list):
    dirs = os.listdir(location)
    for file in dirs:
        #get the date the file was modified/created
        file_mod_time = os.stat(location + "\\" + file).st_mtime
        file_create_time = os.stat(location + "\\" + file).st_ctime

        #mod_date = datetime.datetime.fromtimestamp(os.path.getmtime(location + "\\" + file)).date()
        #create_date = datetime.datetime.fromtimestamp(os.path.getctime(location + "\\" + file)).date()

        #get the date the overview file was modified
        overview_mod_date = os.stat(loc + "\\" + overviewfile + ".xlsx").st_mtime
    
        #only get the files that were modified/added today
        if file_mod_time > overview_mod_date or file_create_time > overview_mod_date:
            #if the file is the overview file, the backup of the overview file, or a master version of a file don't add it to the list
            if not pywildcard.fnmatch(file, overviewfile + ".xlsx") and not pywildcard.fnmatch(file, overviewfile + " *.xlsx") and not file.startswith("~$"):
                #if the file is a directory, go into it and check for new files
                if os.path.isdir(location + "\\" + file):
                    list = find_new_files(location + "\\" + file, list)
                #if the file is an excelsheet, add it to the list
                if pywildcard.fnmatch(file, '*.xlsx'):
                    list.append(location + "\\" + file)
    return list

#append extracted information to the request overview fil
def get_keys_from_init(location, proj_keys, serv_keys, vm_keys):
    section_name = ""
    with open(location) as fp:
        line = fp.readline()
        while line:
            s = line.strip()
            
            # skip blank line
            if not s == "":          
                if section_name == "Projects":

                    if not '[' in s:
                        # extract keys under [Projects] and then continue
                        proj_keys.append(s)                    

                elif section_name == "Services":
                    if not '[' in s:
                        # extract keys under [Services] and then continue
                        serv_keys.append(s)
                
                elif section_name == "VM Cores":
                    if not '[' in s:
                        # extract keys under [VM Cores] and then continue
                        vm_keys.append(s)

                if '[' in s:
                    section_name = check_section(s)
            
            line = fp.readline()
    return proj_keys, serv_keys, vm_keys

def check_section(name):
    print ("\ncheck for section: " + name) 
    
    if name == "[Projects]":
        return "Projects"
    
    elif name == "[Services]":
        return "Services"
    
    elif name == "[VM Cores]":
        return "VM Cores"

def test():
    get_keys_from_init(loc + "/nsoci.ini")

#using the keys found in the config file, pull data from the request forms
def read_from_excel(location, proj_keys, serv_keys, vm_keys, dvm_keys):
    wb = xw.Book(location)
    sht = wb.sheets[0]
    proj_data = {}
    serv_data = {}
    vm_data = {}
    dvm_data = {}
    #get data from the request forms
    for search in proj_keys:
        project_keys = sht.api.UsedRange.Find(search + ":")
        project_values = project_keys.offset(1, 7)
        data = project_values.value
        proj_data[search] = data
    
    for search in serv_keys:
        service_keys = sht.api.UsedRange.Find(search)
        service_values = service_keys.offset(1, 8)

        if service_values.value == "Not to be requested":
            data = None
        else:
            data = service_values.value
        serv_data[search] = data

        if search.startswith("VM") or search.startswith("BM"):
            num_of_cores = service_keys.offset(1, 2)
            print(num_of_cores.address + " " + str(int(num_of_cores.value)))
            for core_num in vm_keys:
                if core_num == str(int(num_of_cores.value)):
                    print("Match found")
                    if data == None:
                        data = 0
                    if core_num in vm_data:
                        vm_data[core_num] = vm_data.get(core_num) + data
                    else:
                        vm_data[core_num] = data


    wb.close()

    return proj_data, serv_data, vm_data, dvm_data

#using the data pulled from the request form, write to the overview file
def write_to_excel(proj_keys, serv_keys, proj_data, serv_data):
    wb = xw.Book(loc + "\\" + overviewfile + ".xlsx")
    sht = wb.sheets[0]
    xlShiftToDown = xw.constants.InsertShiftDirection.xlShiftDown
    current_row = 6
    last_row = 0

    #insert a new row at the bottom of the list
    #currently, the max_rows is 500. This can be changed above if the list grows to exceed 500
    while current_row in range(max_rows):
        if sht.range((current_row, 1)).value == None:
            sht.range((current_row, 1)).api.EntireRow.Insert(Shift=xlShiftToDown)
            print("Row Inserted")
            last_row = current_row - 5
            current_row += max_rows
        current_row += 1
    
    #find the columns that the keys are in and insert the corresponding data into those columns
    for i in range(len(proj_keys)):

        #some of the keys taken from the request file do not match the column titles so I have to hardcode it
        #to look for the correct titles upon coming across those keys
        
        #I was unable to have the overview excel sheet automatically copy the formula to calculate the percentage
        #of 495k and 184k so I have to create it here. the value of four_nine_five_k and one_eight_four_k can be changed above
        if proj_keys[i] == "Total Monthly Cost":
            column_title = sht.api.UsedRange.Find("Budget")
            first_formula_cell = column_title.offset(last_row, 2)
            first_formula_cell.value = '=' + first_formula_cell.offset(1, 0).address + '/' + four_nine_five_k

            second_formula_cell = first_formula_cell.offset(1, 2)
            second_formula_cell.value = '=' + first_formula_cell.offset(1, 0).address + '/' + one_eight_four_k

        elif proj_keys[i] == "Project requestor":
            column_title = sht.api.UsedRange.Find("Resource requestor")

        else:
            column_title = sht.api.UsedRange.Find(proj_keys[i])
        insert_cell = column_title.offset(last_row, 1)
        insert_cell.value = proj_data[proj_keys[i]]

    for i in range (len(serv_keys)):
        column_title = sht.api.UsedRange.Find(serv_keys[i])
        insert_cell = column_title.offset(last_row, 1)
        insert_cell.value = serv_data[serv_keys[i]]

    wb.save()
    wb.close()

def main():
    file_list = []
    project_keys = []
    service_keys = []
    vmcore_keys = []

    project_data = {}
    service_data = {}
    vmcore_data = {}

    #create a list of the new/modified files
    file_list = find_new_files(loc, file_list)
    
    #if files have been added/modified, make a backup of the Total Request Overview
    #excel sheet before proceeding
    if file_list:
        print("new/modified files found")
        make_copy()
    
    #create a list of keys using the section titles of a JIRA tickets. Section titles are 
    project_keys, service_keys, vmcore_keys = get_keys_from_init(loc + "/nsoci.ini", project_keys, service_keys, vmcore_keys)

    print(*file_list, sep = "\n")

    #loop through the list of new files and extract information
    for file in file_list:
        project_data, service_data, vmcore_data, dvmcore_data = read_from_excel(file, project_keys, service_keys, vmcore_keys, dvmcore_keys)
        write_to_excel(project_keys, service_keys, project_data, service_data)

    print(project_data)
    print(service_data)
    print(vmcore_data)

    

if __name__ == "__main__":
    main()