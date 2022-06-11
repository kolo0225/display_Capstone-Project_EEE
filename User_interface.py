###################################################################################################
# Module GUI: user interface to capture the user's inputs to perform the colocation  simulation   # 
#                                                                                                 #
#                                             updated 3/29/21                                     #
#                                                                                                 #
###################################################################################################


import numpy as np
import pandas as pd


from tkinter import Tk, Label, Button, StringVar, OptionMenu, Entry, IntVar, Checkbutton, END, LabelFrame
from tkinter import filedialog


root = Tk()
root.geometry("1920x1080")
root.title("Surplus Interconnection Service")


###########################################################################################
#                                                                                         #
# global variables: utilization_value, year_value, renewable_value, location_value        #
#                                                                                         #
###########################################################################################


#functions
# 2a function for transmission line utilization
def utilization_data():  
    global utilization_value

    utilization_value = 0    
    button_utilization.config(state = 'disabled')
    utilization_input.config(state = 'disabled')
         
    utilization = 0    
    utilization = float(utilization_input.get())
    if utilization > 1:
        utilization = [utilization/100]                 # variable stores the percentage to simulate in a list       
    
    else:
        utilization = [float(utilization_input.get())]  # variable stores the percentage to simulate as list
    
    utilization_value = utilization    
    #print("utilization factor:", utilization)


# 3a function for the period to simulate the renewable resources        
def year_data():
    global year_value  
    
    year_value = 0
    button_years.config(state = 'disabled')
    years_input.config(state = 'disabled')
        
    year = 0
    year = [int(years_input.get())]                     # variable stores years to simulate as list
    #print("years to simulate: ",year)
    year_value = year
   

#5a fuction to let the user to upload a gas forward curve, the code to do the process needs to be added
def gasforwardMLP():
    button_gasforwardMLP.config(state = 'disabled')
    return
#6a fuction to let the user to upload a gas forward curve, the code to do the process needs to be added
def weatherdata():
    button_weatherdata.config(state = 'disabled')
    return
 

#6.5a fuction to let the user to update economic analysis data, the code to do the process needs to be added
def economicanalysis():
    
    button_economicdata.config(state = 'disabled')
    return


# 0 main title
main_title = Label(root, text= "Welcome to Surplus Interconnection Service Simulator", font = 'Arial 14 bold')
main_title.grid(row = 5, column = 1, columnspan = 5, pady = 30, sticky = 'E')


# 1. Select transmission line to simulate based on a list of existing locations with MW nameplate
#drop down list for locations
loc_simulate = Label(root, text= "Select location to simulate: ", font = 'Arial 12')
loc_simulate.grid(row = 13, column = 2, columnspan = 2, padx = 100, pady = 15, sticky = 'W')


options = ['carousel',
           "san_isabel",
           "twin_buttes",
           "colorado_highlands_wind",
           "crossing_trails_wind_farm",
           "kit_carson_wind_farm",
           "new_ts_resource"                     #list of existing renewable resources
]

#Definition of variable location, dropdown menu and set initial location to Carousel
loc = StringVar()                                # variable loc stores selected location as list 
loc.set(options[0])
drop = OptionMenu(root, loc, *options).grid(row = 13, column = 4, pady = 15, sticky = 'W')


# 2 user enters the percentage of transmission line to simulate 
# it calls the function utilization, stores the percentage to simulate in the variable utilization as list
uti_simulate = Label(root, text= "Enter % of transmission line to simulate: ", font = 'Arial 12')
uti_simulate.grid(row = 17,  column = 3, padx = 100, pady = 15, sticky = 'W')


utilization_input = Entry(root, width= 30, borderwidth=5, font = 'Arial 9 bold')
utilization_input.grid(row = 17, column = 4, pady = 15, sticky = 'W')

button_utilization = Button(root, text = "Apply", font = 'Arial 10 bold', command = utilization_data)
button_utilization.grid(row = 17, column = 5, padx = 2, pady =15, sticky = 'W')


    
# 3 user enters how many years to simulate, it calls the function year data
# the number of years is store in the variable year as list
years_simulate = Label(root, text = "Enter how many years to simulate: ", font = 'Arial 12')
years_simulate.grid(row = 21, column = 3, padx = 100, pady = 15, sticky = 'W')

years_input = Entry(root, width= 30, borderwidth = 5, font = 'Arial 9 bold')
years_input.grid(row = 21, column = 4, pady = 15, sticky = 'W')

button_years = Button(root, text = "Apply", font = 'Arial 10 bold', command = year_data)
button_years.grid(row = 21, column = 5, pady =15, sticky = 'W')



#5 the user has the option to update a forward curve or MLP csv file, the code to do the process needs
#the process calls the function gasforwardMLP but does nothing yet
upload_gasforwardMLP = Label(root, text = "Update gas forward or MLP csv file:  ", font = 'Arial 12')
upload_gasforwardMLP.grid(row = 17, column = 6, pady = 15, sticky = 'W')

button_gasforwardMLP = Button(root, text = "Update", font = 'Arial 10 bold', command = gasforwardMLP)
button_gasforwardMLP.grid(row = 17, column = 7, padx = 2, pady =15, sticky = 'W')


#6 the user has the option to update a weather data csv file, the code to do the process needs to be added
#the process calls the function weather data but does nothing yet
upload_weatherdata = Label(root, text = "Update weather data csv file: ", font = 'Arial 12')
upload_weatherdata.grid(row = 13, column = 6,  pady = 15, sticky = 'W')

button_weatherdata = Button(root, text = "Update", font = 'Arial 10 bold', command = weatherdata)
button_weatherdata.grid(row = 13, column = 7, padx = 2, pady =15, sticky = 'W')

#6.5 the user has the option to update economic analysis csv file, the code to do the process needs to be added
#the process calls the function weather data but does nothing yet
upload_economicdata = Label(root, text = "Update economic analysis data csv file: ", font = 'Arial 12')
upload_economicdata.grid(row = 21, column = 6,  pady = 15, sticky = 'W')

button_economicdata = Button(root, text = "Update", font = 'Arial 10 bold', command = economicanalysis)
button_economicdata.grid(row = 21, column = 7, padx = 2, pady =15, sticky = 'W')



#7 the user is given 20 different options of stand-alone or combination of resources to simulate
#the selected values are store in the renewable variable as list
renewable_simulate = Label(root, text = "Select renewable resources to simulate: ", font = 'Arial 11 bold')
renewable_simulate.grid(row = 27, column = 3, columnspan=5, padx = 100, pady = 15, sticky = 'W')




#7a function selected that stores the values of the renewable resources selected to be simulated and cleans screen values

def selected():
    global renewable_value, location_value, plot_value, renewable_value, location_value
    global plot_value, size1_value, size2_value, size3_value, aro_value, batteries_value 
    global CC_value, CT_value, RIC_value, solar_value, wind_value
    
    #myLabel = Label(root, text = loc.get()).grid(row = 9, column = 5, padx = 50, pady = 15, sticky = 'E')
    myButton.config(state = 'disabled')
    
   
      
         
    location = [loc.get()]
    location_value = location    
    #print("location to simulate:", location)
    
    
##################################################################################################      
    
    aro_result = [aro_selection.get()] 
            
    aro_value = aro_result
    #print("Aro :", aro_value)   
   
    batteries_result = [batteries_selection.get()]    
    batteries_value = batteries_result
    #print("Batteries :", batteries_value)  
   
    CC_result = [CC_selection.get()]    
    CC_value = CC_result
    #print("CC:", CC_value) 
   
    CT_result = [CT_selection.get()]    
    CT_value = CT_result
    #print("CT :", CT_value)
   
    RIC_result = [RIC_selection.get()]    
    RIC_value = RIC_result
    #print("RIC :", RIC_value)
   
    solar_result = [solar_selection.get()]    
    solar_value = solar_result
    #print("Solar :", solar_value)
   
    wind_result = [wind_selection.get()]    
    wind_value = wind_result
    #print("Wind :", wind_value)
   
    
    size1_selected = [size1.get(), x_selected.get()]    
    size1_value = size1_selected
    #print("size option #1 x value:", size1_value)

    size2_selected = [size2.get(), y_selected.get()]
    size2_value = size2_selected    
    #print("size option #2 y value:", size2_value)

    size3_selected = [size3.get(), z_selected.get()]
    size3_value = size3_selected    
    #print("size option #3 z value:", size3_value)
  
    selected_p = [selected_plot.get()]
    plot_value = selected_p    
    #print("plot size type:", selected_p)


##############################################################################################################################
   


#8 Button to submit selected renewables resources, it calls the function selected 
myButton = Button(root, text = "Submit", font ='Arial 11 bold', command = selected)
myButton.grid(row = 60, column = 4, columnspan = 2, padx = 130, pady = 25, sticky = 'W')

##############################################################################################################################

#renewable resources
aro_renewable = Label(root, text = "Aro", font = 'Arial 11')
aro_renewable.grid(row = 29, column = 3, columnspan = 2, padx = 160, pady = 6, sticky = 'W')

aro_option = ['No', 'Yes']

aro_selection = StringVar() 
aro_selection.set(aro_option[0])
drop = OptionMenu(root, aro_selection, *aro_option).grid(row = 29, column = 3, padx = 160, pady = 6, sticky = 'E')



batteries_renewable = Label(root, text = "Batteries", font = 'Arial 11')
batteries_renewable.grid(row = 30, column = 3, columnspan = 2, padx = 160, pady = 6, sticky = 'W')

batteries_option = ['No', 'Yes']

batteries_selection = StringVar() 
batteries_selection.set(batteries_option[0])
drop = OptionMenu(root, batteries_selection, *batteries_option).grid(row = 30, column = 3, padx = 160, pady = 6, sticky = 'E')



CC_renewable = Label(root, text = "CC", font = 'Arial 11')
CC_renewable.grid(row = 31, column = 3, columnspan = 2, padx = 160, pady = 6, sticky = 'W')

CC_option = ['No', 'Yes']

CC_selection = StringVar() 
CC_selection.set(CC_option[0])
drop = OptionMenu(root, CC_selection, *CC_option).grid(row = 31, column = 3, padx =160, pady = 6, sticky = 'E')



CT_renewable = Label(root, text = "CT", font = 'Arial 11')
CT_renewable.grid(row = 32, column = 3, columnspan = 2, padx = 160, pady = 6, sticky = 'W')

CT_option = ['No', 'Yes']

CT_selection = StringVar() 
CT_selection.set(CC_option[0])
drop = OptionMenu(root, CT_selection, *CT_option).grid(row = 32, column = 3, padx = 160, pady = 6, sticky = 'E')


RIC_renewable = Label(root, text = "RIC", font = 'Arial 11')
RIC_renewable.grid(row = 33, column = 3, columnspan = 2, padx = 160, pady = 6, sticky = 'W')

RIC_option = ['No', 'Yes']

RIC_selection = StringVar() 
RIC_selection.set(CC_option[0])
drop = OptionMenu(root, RIC_selection, *RIC_option).grid(row = 33, column = 3, padx = 160, pady = 6, sticky = 'E')



solar_renewable = Label(root, text = "Solar", font = 'Arial 11')
solar_renewable.grid(row = 34, column = 3, columnspan = 2, padx = 160, pady = 6, sticky = 'W')

solar_option = ['No', 'Yes']

solar_selection = StringVar() 
solar_selection.set(solar_option[0])
drop = OptionMenu(root, solar_selection, *solar_option).grid(row = 34, column = 3, padx = 160, pady = 6, sticky = 'E')


wind_renewable = Label(root, text = "Wind", font = 'Arial 11')
wind_renewable.grid(row = 35, column = 3, columnspan = 2, padx = 160, pady = 6, sticky = 'W')

wind_option = ['No', 'Yes']

wind_selection = StringVar() 
wind_selection.set(wind_option[0])
drop = OptionMenu(root, wind_selection, *wind_option).grid(row = 35, column = 3, padx = 160, pady = 6, sticky = 'E')



##########################################################################################################################


# 8.1 size choice 1

size_simulate = Label(root, text = "Select farm size to simulate: ", font = 'Arial 11 bold')
size_simulate.grid(row = 27, column = 4, columnspan=5, padx = 8, pady = 15, sticky = 'W')



size1_options = ['Choose a size option #1', 'Additional resource 1', "CC1", "CT1", 
                "RIC1", "ARO1", 'Battery 1']

size1 = StringVar() 
size1.set(size1_options[0])
drop = OptionMenu(root, size1, *size1_options).grid(row = 29, column = 4, padx = 1, pady = 0, sticky = 'W')

x_options = ['0','0.1', '0.2', '0.3', '0.4', '0.5', '0.6', '0.7', '0.8', '0.9', '1']
x_selected = StringVar()
x_selected.set(x_options[0])
drop = OptionMenu(root, x_selected, *x_options).grid(row = 29, column = 4, padx = 30, pady = 0, sticky = 'E')



# 8.2 size choice 2
size2_options = ['Choose a size option #2', 'Additional resource 2', "CC2", "CT2", 
                "RIC2", "ARO2", 'Battery 2']

size2 = StringVar()                                 
size2.set(size2_options[0])
drop = OptionMenu(root, size2, *size2_options).grid(row = 31, column = 4, padx = 1, pady = 0, sticky = 'W')

y_options = ['0', '0.1', '0.2', '0.3', '0.4', '0.5', '0.6', '0.7', '0.8', '0.9', '1']
y_selected = StringVar()
y_selected.set(x_options[0])
drop = OptionMenu(root, y_selected, *y_options).grid(row = 31, column = 4, padx = 30, pady = 0, sticky = 'E')



# 8.3 size choice 3
size3_options = ['Choose a size option #3','Battery 3']

size3 = StringVar()                                 
size3.set(size3_options[0])
drop = OptionMenu(root, size3, *size3_options).grid(row = 33, column = 4, padx = 1, pady = 0, sticky = 'W')

z_options = ['0','0.1', '0.2', '0.3', '0.4', '0.5', '0.6', '0.7', '0.8', '0.9', '1']
z_selected = StringVar()
z_selected.set(z_options[0])
drop = OptionMenu(root, z_selected, *z_options).grid(row = 33, column = 4, padx = 30, pady = 0, sticky = 'E')




##############################################################################################################################


#8.5 Dropdown for plot to show plot
plot_label = Label(root, text= "Select plot for farm size: ", font = 'Arial 11 bold')
plot_label.grid(row = 27, column = 6, columnspan = 1, padx = 30, pady = 20, sticky = 'W')

plot_type = Label(root, text = "Plot type: ", font = 'Arial 11')
plot_type.grid(row = 29, column = 6, columnspan = 2, padx = 30, pady = 6, sticky = 'W')


plot_options = ['Size vs Size','Energy vs Size', "Curtail vs Size", "Utilization vs Size", 
                "1st Derivative vs Size", "2nd Derivative vs Size"]

#Definition of variable location, dropdown menu and set initial location to Carousel
selected_plot = StringVar()                                # variable loc stores selected location as list 
selected_plot.set(plot_options[0])
drop = OptionMenu(root, selected_plot, *plot_options).grid(row = 29, column = 6, padx = 30, pady = 1, sticky = 'E')






#9 next trial
def another_trial():
     # Delete submitted utilization, years and factor information in the entry boxes
        
        
    loc.set(options[0])
    button_utilization.config(state = 'normal')
    button_years.config(state = 'normal')

    
    button_gasforwardMLP.config(state = 'normal')
    button_weatherdata.config(state = 'normal')
    button_economicdata.config(state = 'normal')

    myButton.config(state = 'normal')
    
    utilization_input.config(state = 'normal')
    utilization_input.delete(0, END)
    
    years_input.config(state = 'normal')
    years_input.delete(0, END)
            
     

    
 

try_again_Button = Button(root, text = "Try again", font ='Arial 11 bold', command = another_trial)
try_again_Button.grid(row = 60, column = 6, columnspan = 2, padx = 130, pady = 25, sticky = 'W')


#butt_quit = Button(root, text = 'Quit', font = 'Arial 11 bold', command = root.quit)         # creates quit button & calls command to kill main root widgets     
#butt_quit.grid(row = 60, column = 7, padx = 130, columnspan = 2,  pady = 20, sticky = 'E')


root.mainloop()


