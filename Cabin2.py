import os
from pylab import *
import xlwt
import pickle
import pandas as pd
from datetime import datetime, timedelta
from random import choice, shuffle
from pathlib import Path

    # ### NB: Date Format is Always YYYY-MM-DD, HH:MM ###

def main():
    
    # EDIT HERE ##############################


    # Change this to set the timeframe for bookings

    timeframe_end = '2024-06-05 16:00' 


    # Enter the Date of the Last lottery and The deadline for Applications here:
    last_lottery = '2021-04-20 16:00'
    deadline = '2024-03-26 16:00'

    test = False # When actually running a Lottery, Make sure test = False

    status_message = 'open for booking' # Change this to easily project the message on unbooked days


    ###############################################

    clear_FR_list = False

    status_message = status_message.upper()


    # Building all the Paths 
    applications_filename = r'Cabin Group Signup 1.xlsx'
    former_residents_filename = r'Former Residents.xlsx'

    Winner_file = 'Winner File.pickle'
    Visitor_file = 'Visitors File.pickle'
    
    Student_list_file = 'StudentList.xlsx'


    TimeStamp_string = datetime.strftime(datetime.today(), '%H-%M-%S')
    History_FileName = f'{TimeStamp_string}.xlsx'

    forms_dir = Path('./sample_data')
    Former_dir = Path(forms_dir, 'Former Residents')
    result_dir = Path(forms_dir, 'results')
    result_filename = 'Cabin Winners.xlsx'
    Former_filename = former_residents_filename
    result_path = Path(result_dir, result_filename)
    Former_path = Path(Former_dir, Former_filename)
    

    History_dir = Path(Former_dir, 'History')
    History_path = Path(History_dir, History_FileName)
    
    
    complete_student_list_dir = Path(Former_dir, 'Student Numbers')
    complete_student_list_path = Path(complete_student_list_dir, Student_list_file)

    
    
    former_residents_path = Former_path

    

    applications_path = Path(forms_dir, applications_filename)

    today_string = datetime.strftime(datetime.today(), '%Y-%m-%d')
    result_filename = f'{today_string}_result.xls'

    result_path = Path(result_dir, result_filename)

    deadline = datetime.strptime(deadline, '%Y-%m-%d %H:%M')
    lasttime = datetime.strptime(last_lottery, '%Y-%m-%d %H:%M')
    today_string = datetime.strftime(datetime.today(), '%Y-%m-%d')
    timeframe_end = datetime.strptime(timeframe_end, '%Y-%m-%d %H:%M')
    timeframe_start = deadline
    

    # If a Test is conducted the results are saved in a test file instead

    if test != False:
        Winner_file = 'Test File.pickle'
        Visitor_file = 'Visitor File Test.pickle'
        print('!!!TEST RUN, Former Residents.xlsx will not be updated!!!'.upper())

    # if the directory for the results does not exist, make it
    if not os.path.isdir(result_path.parent):
        os.mkdir(result_path.parent)

    if not os.path.isdir(Former_path.parent):
        os.mkdir(Former_path.parent)
        Former_path = Path(forms_dir, Former_filename)
        former_residents_path = Former_path
        os.mkdir(complete_student_list_path.parent)
        os.mkdir(History_path.parent)
        

    if not os.path.isdir(History_path.parent):
        os.mkdir(History_path.parent)
    if not os.path.isdir(complete_student_list_path.parent):
        os.mkdir(complete_student_list_path.parent)
    

    for path in [applications_path]:
        if not os.path.isfile(path):
            raise ValueError(f'{path} does not exist. Check input files.')
        
    
    
    # Reads the actual applications:
    applications = pd.read_excel(
                applications_path,
                usecols=['Navn','Terms and Conditions', 'E-postadresse','Fullføringstidspunkt' ,'Your Student ID', 'Student ID of the other students', 'Date from', 'Date to', 'Details'],
                engine='openpyxl'
                )


    # Remove the timezone data from the applications, so they can be compared
    # to the deadline
    application_times = [t.replace(tzinfo=None) for t in applications['Fullføringstidspunkt']]
    before_deadline = [t < deadline for t in application_times]
    after_last = [t > lasttime for t in application_times]
    keep = [a and b for a,b in zip(before_deadline, after_last)]
    applications = applications[keep]

    application_times_1 = [t.replace(tzinfo=None) for t in applications['Date to']]
    application_times_2 = [t.replace(tzinfo=None) for t in applications['Date from']]

    before_timeframe_end = [t < timeframe_end for t in application_times_1]
    after_timeframe_start = [t > timeframe_start for t in application_times_2]
    keep = [a and b for a,b in zip(before_timeframe_end, after_timeframe_start)]
    applications = applications[keep]

    if applications.empty == True:
        return print('ERROR, NO APPLICATIONS IN THIS PERIOD')
    
    
    applications = applications[applications['Terms and Conditions'].notna()]

   
    applications.sort_values(by = ['Date from'], inplace=True)
    applications['overlap'] = (applications['Date to'].shift()-applications['Date from']) > pd.Timedelta(0)
    applications = applications.loc[applications['Terms and Conditions'] != 'No']

    
    # Checks if The file for former Residents Exist
    if os.path.isfile(Former_path):
        
        # Opens up the excel list with former residents
        Former_Residents_df = pd.read_excel(
            former_residents_path,
            usecols=['Former Residents'],
            engine='openpyxl'
            )
        Former_residents = list(Former_Residents_df['Former Residents'])

    # if it doesent exist a new file is made
    if not os.path.isfile(Former_path):
        Former_residents = []

    # Moves the files to the right position

    if os.path.isfile(Path(forms_dir, Former_filename)):
        os.remove(Path(forms_dir, Former_filename))

    Former_path = Path(Former_dir, Former_filename)
    former_residents_path = Former_path
    
    # running the applications trough a set of filters
    # These filters gives each group a priority score based on a set of parameters and works by modifying the dataframe with an updated score
   

    Number_of_nights(applications) # Adds the column "Number of Nights"
    applications = Student_ID_Validation(applications, complete_student_list_path, Student_list_file, forms_dir) # Filters fake Student ID's
    Visitor_check(applications, Former_residents) # Counts the number of visitors in the bookings
    applications, week = date_check(applications) 
    priority_eval_size(applications)
    applications = priority_eval_stays(applications)
    first_time_goers(applications)
    add_ID(applications)
    Weekend_priority(applications)

    

    # filters out smartasses that tries to apply with the same group several times but with different applicants
    applications = filter_duplicates(applications)

    ##############################This is where the fun starts##############################

    # formates the bookings into a calendar-like dataframe to check for overlap

    days = pd.date_range(min(list(applications['Date from'])), max(list(applications['Date to']+timedelta(days=1)))-timedelta(days=1), freq='d')

    days_df = pd.DataFrame(columns=days)
    BetweenDates = [min(days), max(days)]



    booking_IDs = list(applications['Application ID'])
    day_from = list(applications['Date from'])
    
    range_dict = dict(zip(applications['Date from'], applications['Date to']))
    daterange = list(range(len(days)))
    Id_to_student_Id_dict = dict(zip(applications['Application ID'], applications['Your Student ID']))
    
   # Creates an enourmus dataframe consisting of overlapping bookings and NaN's
    n = 0
    
    
    for day in day_from:
        
        
        i = 0
        new_row = array(['NaN'] * len(days))
        while i < len(daterange):
            
            
            if day == days_df.columns[i]:
                new_row[i] = booking_IDs[n]
                new_index = i
                
                day_count = 0
                
                while new_index < len(daterange):
                    new_row[new_index] = booking_IDs[n]
                    day_count += 1
                    
                    if days_df.columns[new_index] == range_dict[day]:
                        day_count = 0
                        break
                    if day_count == 3:
                        B_id = booking_IDs[n]
                        
                        print(f'Break for booking {Id_to_student_Id_dict[B_id]}, booked from {day} to {range_dict[day]}')
                        break
                    
                    new_index += 1
            i += 1
        days_df.loc[len(days_df)] = new_row
        
        n = n + 1
    
    
    
    result_list = []
    group_ID_result = []
    GROUP_ID_DICT = dict(zip(applications['Application ID'], applications['Group ID']))
    Dict_id = ID_dict(applications, status_message)
   
    # converts the dataframe into a list of lists and removes NaN's

    
    Filtered_application_list = filter_nan(days_df, applications, status_message)
    
    
    # Filters out applications with the lower score
    # Adds the status message on days that haven't been booked
    # if two or more bookings has the same priority-score there is a lottery
    for application_list in Filtered_application_list:
        winner = do_lottery(application_list, applications, status_message)
        if winner != status_message:
            if GROUP_ID_DICT[winner] in group_ID_result:
                application_list = application_list.remove(int(winner))
                winner = do_lottery(application_list, applications, status_message)
            group_ID_result.append(int(GROUP_ID_DICT[winner]))
        result_list.append(winner)
    result_list = [Dict_id[i] for i in result_list]

    result_list = Filter_result(result_list, status_message)
    # The result is converted into a dataframe
    result_df = pd.DataFrame(zip(days, result_list))
    result_df.columns =['Dates', 'Winners']
    result_dict = dict(zip(result_df['Dates'], result_df['Winners']))

    group_winner_dict = dict(zip(applications['Your Student ID'], applications['Duplicates']))
    
    # Winners are added to the list of former residents
    Former_residents_and_winners = add_list(Former_residents, winner_list(result_list, status_message))
    Winner_group = list(applications['Duplicates'])

    W_list = winner_list(result_list, status_message)
   
    Winner_group = [group_winner_dict[int(i)] for i in W_list]
    Winner_group = flatten_extend(Winner_group)
    
    Former_residents_and_winners = add_list(Former_residents_and_winners, Winner_group)
    
    
    if clear_FR_list == True:
        Former_residents_and_winners = ['']
    # The list with Former residents and winners are converted into a dataframe
    # The dataframe is also stored as both an excel spreadsheet and csv file
    df_2 = pd.DataFrame(Former_residents_and_winners)
    df_2.columns = ['Former Residents']
    

    # Updates the list over former residents if the script is run for real
    if test == False:
        df_2.to_excel(Former_path)
        df_2.to_excel(History_path) # Keeps track of history just in case
        
        print(f'Former Residents updated at {datetime.now()}')
       
    
    # The winners and former cabin-goers are pickled
    with open(Path(result_dir, Winner_file), 'wb') as fp:
        pickle.dump(result_dict, fp)

    with open(Path(Former_dir, Visitor_file), 'wb') as fp:
        pickle.dump(Former_residents_and_winners, fp)
    
   
   

    # Bookings are converted to an excel spreadsheet here
        
    write_to_excel(['Winner of The cabin'], [result_dict], result_path, BetweenDates, week, applications, status_message)

    
   
    
    print(f'Written results to {result_path}')
    print(f'Written at {datetime.now()}')
    print('Done.')

# Checks the number of nights
def Number_of_nights(applications):
    test = False
    if test == True:
        applications['Date from'] = pd.to_datetime(applications['Date from'], format='%D-%M-%Y')
        applications['Date to'] = pd.to_datetime(applications['Date to'], format = '%D-%M-%Y')
        delta = applications['Date to'] - applications['Date from']
        series = pd.Series(delta)
        date_to = applications['Date to'].dt.day.values
        date_from = applications['Date from'].dt.day.values
        num_nights = date_to - date_from
        applications['Number of Nights'] = num_nights

    if test == False:
        num_nights = []
        day_from = list(applications['Date from'])
        day_to = list(applications['Date to'])
        i = 0
        while i < len(day_from):
            date_range = pd.date_range(day_from[i], day_to[i])
            num_nights.append(len(date_range))
            i += 1
        applications['Number of Nights'] = num_nights

# Checks how many former residents are on the application
# Also Assess the group size
def Visitor_check(applications, Former_residents):
    former_count_list = []
    your_student_id = list(applications['Your Student ID'])
    index = 0
    student_group_list = [eval(i) for i in list(applications['Student ID of the other students'])]
    group_size = []
    size = 0
    for student_group in student_group_list:
        size = len(student_group)
        if your_student_id[index] not in student_group:
            size = size + 1
        group_size.append(size)
        former_count = 0
        former_count = len([i for i in student_group if i in Former_residents])
        if your_student_id[index] in Former_residents:
            former_count = former_count + 1
        index = index + 1
        former_count_list.append(former_count)
    applications['Total Group Size'] = group_size
    applications['Number of Former Visitors'] = former_count_list

# checks what weekdays people are booking at
    # Useful if there is different parameters added later
def date_check(applications):
    days = pd.date_range(min(list(applications['Date from'])), max(list(applications['Date to']+timedelta(days=1)))-timedelta(days=1), freq='d')
    
    
    date_dict = {0 : "Monday", 
                 1 : "Tuesday", 
                 2 : "Wednesday", 
                 3 : "Thursday", 
                 4 : "Friday", 
                 5 : "Saturday", 
                 6 : "Sunday"} 
    date_dict_2 = {0 : "Weekday", # Monday
                 1 : "Weekday", # Tuesday
                 2 : "Weekday", # Wednesday
                 3 : "Weekday", # Thursday
                 4 : "Weekend", # Friday
                 5 : "Weekend", # Saturday
                 6 : "Weekend"} # Sunday
    
    week = [date_dict[datetime.weekday(i)] for i in days]

    Start_day = []
    End_day = []

    Start_day_stat = []
    End_day_stat = []
    i = 0
    daterange_from = pd.date_range(min(list(applications['Date from'])), max(list(applications['Date from']+timedelta(days=1)))-timedelta(days=1), freq='d')
    daterange_to = pd.date_range(min(list(applications['Date to'])), max(list(applications['Date to']+timedelta(days=1)))-timedelta(days=1), freq='d')
    while i < len(applications['Date from']):
        Start_day.append(date_dict[datetime.weekday(daterange_from[i])])
        Start_day_stat.append(date_dict_2[datetime.weekday(daterange_from[i])])
        End_day.append(date_dict[datetime.weekday(daterange_to[i])])
        End_day_stat.append(date_dict_2[datetime.weekday(daterange_to[i])])
        i = i + 1
    
    

    applications['Arriving Day'] = Start_day
    applications['Departure Day'] = End_day

    applications['A_day_stat'] = Start_day_stat
    applications['D_day_stat'] = End_day_stat


    return applications, week

# adds priority score based on group size
def priority_eval_size(applications):
    priority_array = zeros(len(applications['Navn']))
    group_size = list(applications['Total Group Size'])
    i = 0
    for group in group_size:
        if group >= 8:
            priority_array[i] = priority_array[i] + 10
        i = i + 1
    applications['Priority Score'] = priority_array

# adds priority score based on number of nights
def priority_eval_stays(applications):
    priority_array = array(applications['Priority Score'])
    number_of_days = list(applications['Number of Nights'])
    i = 0
    for days in number_of_days:
        if days >= 3 and days < 5:
            priority_array[i] = priority_array[i] + 5
        # Bookings exceeding 4 nights will result in removal of priority score
        # This is to prevent people of taking the advantage of the system
        if days >= 5:
            priority_array[i] = priority_array[i] - 20
        if priority_array[i] < 0:
            priority_array[i] = priority_array[i] 
        # If some people get the idea of booking for more than a week their application will be removed
        # this is filter is hopefully not nessescary
        if days > 15:
            priority_array[i] = 'REMOVE'
        i = i + 1
    applications['Priority Score'] = priority_array
    applications = applications.drop(applications[applications['Priority Score'] == 'REMOVE'].index)
    return applications

# adds priority score based on how many in the group aren't found in the list over former visitors
def first_time_goers(applications):
    visitors = list(applications['Number of Former Visitors'])
    Priority_array = array(applications['Priority Score'])
    i = 0
   
    for person in visitors:
        if person == 0:
            Priority_array[i] = Priority_array[i] + 20  # All priority scores can be changed if needed
        i = i + 1
    applications['Priority Score'] = Priority_array

# Adds a booking ID to each group
def add_ID(applications):
    column_length = len(applications['Navn'])
    ID = arange(0, column_length)
    applications['Application ID'] = ID

# sets of dictionaries
def ID_dict(applications, status):
    application_id, your_id = list(applications['Application ID']), list(applications['Your Student ID'])
    application_id.append(status), your_id.append(status)
    Dict_id = dict(zip(application_id, your_id))
    return Dict_id


def score_dict(applications):
    dict_df = dict(zip(applications['Application ID'], applications['Priority Score']))
    return dict_df

# filters out NaN's and converts the dataframe to a list of lists
def filter_nan(days_df, applications, status_message):
    days = pd.date_range(min(list(applications['Date from'])), max(list(applications['Date to']+timedelta(days=1)))-timedelta(days=1), freq='d')

    application_list = []
    for ColNames in days_df.columns:
        application_list.append(list(days_df[ColNames]))

    Filtered_application_list = []
    for dates in application_list:
        new_list = []
        i = 0
        while i < len(dates):
            if dates[i] != 'NaN':
                new_list.append(dates[i])
            i += 1
        if len(new_list) == 0:
            new_list.append(status_message)
        Filtered_application_list.append(new_list)
    
    return Filtered_application_list

# The actual lottery
# Bookings are selected based on highest priority score
# if several bookings have the same score a lottery is conducted
# Days without bookings are marked as open
def do_lottery(application_list, applications, status_message):
    group_ID_dict = dict(zip(applications['Application ID'], applications['Group ID']))
    Back_to_A_ID = dict(zip(applications['Group ID'], applications['Application ID']))

    if len(application_list) == 0:
        print('List empty')
    score = dict(zip(applications['Application ID'], applications['Priority Score']))

    if application_list[0] == status_message:
        result = status_message
        return result
    else:
        test_list = [int(i) for i in application_list]

        big_count = 0
        i = 1
        biggest = int(test_list[0])
        big_index = 0
        score_list = []
        while i < len(test_list):
            score_list.append(score[int(test_list[i])])
            if score[biggest] < score[int(test_list[i])]:
                biggest = int(test_list[i])
                big_index = i
            i += 1
        lottery = []
        for n in test_list:
            if score[int(n)] == score[biggest]:
                lottery.append(group_ID_dict[int(n)]) # Saves the lottery list as group ID

        lottery = set(lottery) # Removes Duplicates for the same dates
        lottery = [Back_to_A_ID[i] for i in lottery] # Converts it back to application ID
        
        shuffle(lottery) # Shuffles to list to reduce potential bias
        result = choice(lottery) # picks one random application ID from the list 
        result = int(result)
        
            
        return result

# Takes the results and removes the unbooked dates
def winner_list(result_list, status_message):
    winners = []
    for result in result_list:
        if result != status_message:
            winners.append(result)
    return winners

# Merges two lists
def add_list(old_entities, new_entities):
    for entity in new_entities:
        if entity not in old_entities:
            old_entities.append(entity)
    return old_entities

# The function that filters out smartass-attempts
# Sorts each group and check if there are duplicates
# Assign each group a group ID 
# group ID's are used to filter out double-bookings by the same group on same dates
# also used to prevent some groups to get several bookings
def filter_duplicates(applications):

    Duplicate_drop = False # change to True to limit each group to one booking

    student_group_list = [eval(i) for i in list(applications['Student ID of the other students'])]
    applicants_id = list(applications['Your Student ID'])
    index = 0
    Group_list = []
    while index < len(student_group_list):
        groups = list(student_group_list[index])
        groups.append(applicants_id[index])
        sort(groups)
        groups = tuple(groups)
        Group_list.append(groups)
        index = index + 1
    
    applications['Duplicates'] = Group_list
    applications.sort_values(by=['Duplicates'])
    if Duplicate_drop == True:
        applications.drop_duplicates('Duplicates', keep = 'last', inplace=True)
        applications = applications.drop(columns='Duplicates')

    Group_ID = array(['NaN'] * len(Group_list))
    ID_index = 0
    for group in Group_list:
        i  = 0
        while i < len(Group_list):
            if group == Group_list[i]:
                Group_ID[i] = ID_index
            i += 1
        ID_index += 1
    if 'NaN' in Group_ID:
        raise ValueError('Uncomplete Loop')
    applications['Group ID'] = Group_ID
    return applications

# converts the result into an excel spreadsheet
def write_to_excel(sheet_names, winnerdicts, result_path, BetweenDates, week, applications, status_message):

    assert len(sheet_names) == len(winnerdicts), 'you need one sheet name for every winner dict'

    Name_dict = dict(zip(applications['Your Student ID'], applications['Navn']))
    Name_dict[status_message] = ''


    wb = xlwt.Workbook() 
    line_width = 20

    style_header_container = xlwt.easyxf("alignment: wrap True; font: bold on, height 300; borders: top thin, bottom thin, left thin, right thin")
    style_header           = xlwt.easyxf("alignment: wrap True; borders: left thin, right thin, top thin, bottom thin; font: bold on")
    style_header_2           = xlwt.easyxf("alignment: wrap True; font: bold on, height 250; borders: top thin, bottom thin, left thin, right thin")
    style                  = xlwt.easyxf("alignment: wrap True, vert centre; borders: left thin, right thin, top thin, bottom thin; font: bold on")
    From_date, To_date = datetime.strftime(BetweenDates[0], '%Y-%m-%d'), datetime.strftime(BetweenDates[1], '%Y-%m-%d')
    # create the sheets
    sheetlist = [wb.add_sheet(name) for name in sheet_names]

    today_string = datetime.strftime(datetime.today(), '%Y-%m-%d')
    style1 = xlwt.easyxf("alignment: wrap True, vert centre; borders: left thin, right thin, top thin, bottom thin; font: bold on",num_format_str='MM-DD-YY')

    for sheet, result, header in zip(sheetlist, winnerdicts, sheet_names):
        # set size for columns
        sheet.col(0).width = 200 * line_width + 1000
        sheet.col(1).width = 200 * line_width + 2000

        sheet.col(2).width = 200 * line_width + 3000
        sheet.col(3).width = 200 * line_width + 2000

        

        sheet.write_merge(0, 0, 0, 3, f'{header} {today_string}', style_header_container)

        sheet.row(0).height_mismatch = True       # for the adjustment of the row height
        sheet.row(0).height = 1000
        
        sheet.write_merge(1, 2, 0, 3, f'From {From_date} to {To_date}', style_header_2)
        # write header
        sheet.write(3, 0, 'Day', style_header)
        sheet.write(3, 1, 'Date', style_header)
        sheet.write(3, 2, 'Student ID',style_header)
        sheet.write(3, 3, 'Name', style_header)
        

        row = 4 # start row
        i = 0
        for date, Winner in result.items():
            # separate items by linebreak
            
            sheet.write(row, 0, week[i], style)
            sheet.write(row, 1, date, style1)
            sheet.write(row, 2, Winner, style)
            sheet.write(row, 3, Name_dict[Winner], style)
            
            sheet.row(row).height_mismatch = True
            row = row + 1
            i += 1

    wb.save(result_path)

# ###################This function doesent work yet###################################
def Inspice_Student_ID(applications):
    Email_list = list(applications['E-postadresse'])
    student_dict = dict(zip(list(applications['Your Student ID']), list(applications['E-postadresse'])))
    ID_check_dict = {}
    for Email, Student_ID in student_dict.items():
        Mail = Email[:6]
        if Mail != Student_ID:
            ID_check_dict = {Student_ID : 'Fake ID'}
        else:
            ID_check_dict = {Student_ID : 'Real ID'}
    return ID_check_dict
# ######################################################################################

# takes a list of list and converts it to a single list
def flatten_extend(matrix):
    flat_list = []
   
    for row in matrix:
        row = list(row)
        flat_list.extend(row)
    return flat_list

# ################# In Progress ##########################
def ID_TO_NAMES(applications, result_dict, status_message):
    result_winner = []
    result_dates = []
    for dates, winner in result_dict.items():
        result_dates.append(dates)
        result_winner.append(winner)
    Name_dict = dict(zip(applications['Your Student ID'], applications['Navn']))
    Name_dict[status_message] = status_message
    result_names = [Name_dict[i] for i in result_winner]
    result_dict = dict(zip(result_winner, result_names))
    Nested_Name_dict = {}
    Nested_dict = {idx: {key : result_dict[key]} for idx, key in zip(result_dates, result_dict)}
    return Nested_dict
# ############################################################

# Removes one-day bookings that gets granted
# since it means the group doesen't get a night at the cabin
def Filter_result(result_list, status_message):
    result_array = array(result_list)
    i = 1
    if result_array[0] != result_array[1]:
        result_array[0] = status_message

    while i < len(result_array) - 1:
        if result_array[i] != status_message:
            if result_array[i] != result_array[i-1] and result_array[i] != result_array[i+1]:
                result_array[i] = status_message
        i += 1
    result_list = []
    for n in result_array:
        if n != status_message:
            result_list.append(int(n))
        if n == status_message:
            result_list.append(n)
            
    return list(result_list)

# Gives Priority points to groups that books over a weekend
# Makes sure people are preferably booking friday to sunday
def Weekend_priority(applications):
    weekday_dict_from = dict(zip(applications['Application ID'], applications['A_day_stat']))
    weekday_dict_to = dict(zip(applications['Application ID'], applications['D_day_stat']))
    A_id = list(applications['Application ID'])
    Priority_score = array(applications['Priority Score'])
    i = 0
    while i < len(A_id):
        if weekday_dict_from[A_id[i]] == 'Weekend' or weekday_dict_to[A_id[i]] == 'Weekend':
            Priority_score[i] = Priority_score[i] + 5
            if weekday_dict_from[A_id[i]] == 'Weekend' and weekday_dict_to[A_id[i]] == 'Weekend':
                Priority_score[i] = Priority_score[i] + 20
        i += 1
    applications['Priority Score'] = Priority_score


# imports a list over actual student numbers and remove fake ones

def Student_ID_Validation(applications, complete_student_list_path, Student_list_file, forms_dir):
    new_path = Path(forms_dir, Student_list_file)

    if os.path.isfile(complete_student_list_path):
        if os.path.isfile(new_path):
            os.remove(new_path)

        Student_list_file = pd.read_excel(
                complete_student_list_path,
                usecols=['ID'],
                engine='openpyxl'
                )
        Student_list_file = list(Student_list_file['ID'])
        


        New_list = []
        for n in list(applications['Your Student ID']):
            if n not in Student_list_file:
                New_list.append('False')
            else:
                New_list.append(n)
            
        applications['Your Student ID'] = New_list
        mask = applications['Your Student ID'] == 'False'

        # Removes applications where the applicant uses a fake student ID
        applications = applications[~mask]

        group = list(applications['Student ID of the other students'])
        New_list = []

        # Removes the fake student numbers from the rest of the group
        # This keeps the application but makes sure a group doesent get
        # a higher priority than deserved
        for n in group:
            a = [str(i) for i in n if i in Student_list_file]
            New_list.append(str(a))
        applications['Student ID of the other students'] = New_list

        return applications
    
    if os.path.isfile(new_path):

        Student_list_file = pd.read_excel(
                new_path,
                usecols=['ID'],
                engine='openpyxl'
                )
        Student_list_file = list(Student_list_file['ID'])
        


        New_list = []
        for n in list(applications['Your Student ID']):
            if n not in Student_list_file:
                New_list.append('False')
            else:
                New_list.append(n)
            
        applications['Your Student ID'] = New_list
        mask = applications['Your Student ID'] == 'False'
        applications = applications[~mask]

        group = list(applications['Student ID of the other students'])
        New_list = []
        for n in group:
            a = [str(i) for i in n if i in Student_list_file]
            New_list.append(str(a))
        applications['Student ID of the other students'] = New_list

        # if the student list file is placed in /sample_data it gets moved
        # to the right directory
        os.rename(new_path, complete_student_list_path)

        return applications
    
    else:
        print(f'{Student_list_file} not found')
        print(f'Please Place {Student_list_file} in {complete_student_list_path}\n ')
        
        
        return applications
    

if __name__ == "__main__":
    main()
    