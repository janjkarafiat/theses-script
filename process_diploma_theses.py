import csv
from openpyxl import load_workbook


def check_row_content(sheet, cell:str):
    content = sheet[cell].value
    if content == None:
        return False
    else:
        return True


def create_thesis_info(sheet, row_num:int):
    thesis = sheet['A'+str(row_num)].value + ': ' + sheet['C'+str(row_num)].value + '. ' + 'URL of thesis archive: ' + sheet['D'+str(row_num)].value + ' ' + 'Supervisor: ' + sheet['E'+str(row_num)].value + ', Reader: ' + sheet['F'+str(row_num)].value + '.'
    return thesis


def add_thesis_to_database(database:dict, name:str, study_programme:str, thesis:str):
    if database.get(name, 0) == 0:
    # no name, no study_programme, no thesis
        database.setdefault(name, {study_programme: [thesis]})

    elif database[name].get(study_programme, 0) == 0:
    # name exist, no study_programme, no thesis
        database[name].setdefault(study_programme, [thesis])
    
    else:
    # name and study_programme exist, thesis append only
        database[name][study_programme].append(thesis)


def make_sorted_lst_of_names(db:dict):
    my_lst = []
    for item in inventory.keys():
        item = item.strip(' .,;')
        my_lst.append(item)
    result = sorted(my_lst)
    return result


def extract_data_from_csv(csv_file):
    result = []
    with open(csv_file, 'r', newline='', encoding='UTF-8') as file:
        reader = csv.reader(file, delimiter=';')
        for row in reader:
            if row[1] == 'ID':
                continue
            if row == None:
                break
            data = row
            result.append(data)
    return result


def check_name_in_lst(name_tab:str, my_lst:list):
    test = name_tab.split()
    if len(test) >= 3:
        name, surname = test[0], test[1]

    else: 
        name, surname = name_tab.split()

    for index, item in enumerate(my_lst):
        if surname in item[0]:
            if name in item[0]:
                return (1,index)
    else:
        return (0,0)


def make_amount(count_of_values:list,amounts:list):
    a = count_of_values[0] * amounts[0]
    b = count_of_values[1] * amounts[1]
    c = count_of_values[2] * amounts[2]
    d = count_of_values[3] * amounts[3]
    result = [a,b,c,d]
    return sum(result)


def write_data_to_file(db:dict, txt_file):
    for name, data in db.items():
        txt_file.write(name + '\n')
        txt_file.write('\n')
        for programme, theses in data.items():
            txt_file.write(programme + ':' + '\n')
            for thesis in theses:
                txt_file.write(thesis + '\n')
            txt_file.write('\n')
        txt_file.write('X' * 40 + '\n')
        txt_file.write('\n')


# load excell file
wb = load_workbook(filename = "theses.xlsx")

# 9 sheets
AB_BC = wb['AB BC']
AB_MGR = wb['AB MGR']
USU_BC = wb['USU BC']
MIT_BC = wb['MIT BC']
MIT_MGR = wb['MIT MGR']
HSA_BC = wb['HSA BC']
GNM_MGR = wb['GNM MGR']
VEU_MGR = wb['VEU MGR']
DVZ_MGR = wb['DVZ MGR']

ALL_SHEETS = (AB_BC, AB_MGR, USU_BC, MIT_BC, MIT_MGR, HSA_BC, GNM_MGR, VEU_MGR, DVZ_MGR)
lst_of_str_progs = ['AB BC', 'AB MGR', 'USU BC', 'MIT BC', 'MIT MGR', 'HSA BC', 'GNM MGR', 'VEU MGR', 'DVZ MGR']

# main database, dict, {key 'name', value nested dict {key 'str_prog', value [list of theses]}}
inventory = {}

# main for loop, create a whole database
for i, prog in enumerate(ALL_SHEETS):
    str_prog = lst_of_str_progs[i]

    # supervisor loop, column E
    number_of_row = 3
    supervisor_loop = True
    while supervisor_loop:
        cell = 'E'+ str(number_of_row)
        if check_row_content(prog, cell) == False:
            break
        supervisor = prog[cell].value
        supervisor = supervisor.strip(' .,;')
        thesis = create_thesis_info(prog, number_of_row)
        add_thesis_to_database(inventory, supervisor, str_prog, thesis)
        number_of_row += 1

    # reader_loop, column F
    number_of_row = 3
    reader_loop = True
    while reader_loop:
        cell = 'F'+ str(number_of_row)
        if check_row_content(prog, cell) == False:
            break
        reader = prog[cell].value
        reader = reader.strip(' .,;')
        thesis = create_thesis_info(prog, number_of_row)
        add_thesis_to_database(inventory, reader, str_prog, thesis)
        number_of_row += 1

print()
print("---------------")
print("DATA PROCESSED.")
print("---------------")


# Write data to txt file
data_file = open('all_theses.txt', mode='w')
write_data_to_file(inventory, data_file)
data_file.close()


print()
print("------------------------------------")
print("DATA WRITTEN TO FILE all_theses.txt")
print("------------------------------------")

# make a statistics
lst_of_names = make_sorted_lst_of_names(inventory)

statistic = {}

for name in lst_of_names:
    statistic[name] = {'BC_SUPERVISOR': 0, 'BC_READER': 0, 'MGR_SUPERVISOR': 0, 'MGR_READER': 0}

BC_PROG = ['AB BC', 'USU BC', 'MIT BC', 'HSA BC']
MGR_PROG = ['AB MGR', 'MIT MGR', 'GNM MGR', 'VEU MGR', 'DVZ MGR']

for name, data in inventory.items():
    for programme, theses in data.items():
        for thesis in theses:
            if programme in BC_PROG:
                if "Supervisor: " + str(name) in thesis:
                    statistic[name]['BC_SUPERVISOR'] += 1
                if "Reader: " + str(name) in thesis:
                    statistic[name]['BC_READER'] += 1
            else:
                if "Supervisor: " + str(name) in thesis:
                    statistic[name]['MGR_SUPERVISOR'] += 1
                if "Reader: " + str(name) in thesis:
                    statistic[name]['MGR_READER'] += 1


emp_list_of_names = extract_data_from_csv("employees.csv")
actual_amounts = extract_data_from_csv("amounts.csv")

amounts = []
for item in actual_amounts:
    amount = int(item[1])
    amounts.append(amount)

final_statistic = []

for item in lst_of_names:
    one_row = ['',]
    one_row[0] = item

    emp_value = check_name_in_lst(item, emp_list_of_names)
    if emp_value[0] == 1:
        one_row[1:4] = emp_list_of_names[emp_value[1]]            
    else:
        one_row[1:4] = ['---', '---', '---']

    one_row[4:8] = list(statistic[item].values())

    one_row.insert(8, sum(one_row[4:8]))

    amount_of_crowns = make_amount(one_row[4:8], amounts)
    one_row.insert(9, str(amount_of_crowns))

    final_statistic.append(one_row)


# write data to new csv file
tab_header = ['NAME','NAME IN EMPLOYEES.CSV','ID','EMPLOYMENT TYPE','BC SUPERVISOR','BC READER','MGR SUPERVISOR','MGR READER','THESES COUNT','PAYMENT']
final_statistic.insert(0,tab_header)

final_csv_file = "statistics.csv"

with open(final_csv_file, mode='w', newline='', encoding='utf-8') as file:
    writer = csv.writer(file, delimiter=';')
    
    writer.writerows(final_statistic)


print()
print("-----------------------------------")
print("DATA WRITTEN TO FILE statistics.csv")
print("-----------------------------------")