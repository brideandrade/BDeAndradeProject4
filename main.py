#Briana DeAndrade
#Nothing was left undone according to the directions

#I had to import the openpyxl and numbers packages in order to open the Excel file as well as do certain math with within the code relating to the excel
import openpyxl
import numbers
import openpyxl.utils

#Here is my main function
    #This is where I opened the worksheet using the open_worksheet fxn and stored it in a variable
    #This is where I also stored the should_get_losses fxn in a variable as requested
    #I passed the two variables as a parameter using the process_data fxn
def main():
    pop_data_sheet = open_worksheet("countyPopChange2020-2021.xlsx")
    show_losses = should_get_losses()
    process_data(pop_data_sheet, show_losses)

#This is my open_worksheet fxn where I "opened" the workbook file in order to read it using the "active" piece
    #I returned the worksheet from the fxn

def open_worksheet(file):
    pop_excel = openpyxl.load_workbook(file)
    data_sheet = pop_excel.active
    return data_sheet

#Here is my should_get_losses fxn
    #This is where I prompt the user the question
        #Depending on what the user responds, the return will differ
def should_get_losses():
    show_losses = input("Should we get the counties that lost population?")
    if show_losses == "yes":
        return True
    else:
        return False

#This is my process_data fxn:
    #Here I take two parameters (data_sheet, and show_losses)
    #The math takes place here where we calculate the % of pop change
        #This relates to the show_losses piece where if show_losses is true, the state, county names and % change... (as in Proj directions) are printed
            #If false, (as in Proj directions) are printed
def process_data(data_sheet, show_losses):
    for row in data_sheet.rows:
        state_cell = row[5]
        county_cell = row[6]
        population_row = row[9]
        pop_row = row[11]
        population_est = population_row.value
        pop_change = pop_row.value
        if not isinstance(pop_change, numbers.Number):
            continue
        population_change_percent = pop_change/population_est
        if show_losses and population_change_percent < -.02:
            print(f"In {state_cell}, {county_cell} had a {population_change_percent}% decrease between July 2020 - July 2021.")
        if show_losses and population_change_percent > .015:
            print(f"In {state_cell}, {county_cell} had a {population_change_percent}% increase between July 2020 - July 2021.")

main()