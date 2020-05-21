from tkinter import *
from PIL import ImageTk, Image
from tkinter import messagebox, filedialog
from openpyxl.styles import Font
import openpyxl, os

# Making the main window that popups at the start
mainWin = Tk()
mainWin.title("Tim's Strength Program v1.0.0 BETA")
mainWin.iconbitmap('barbell.ico')
mainWin.geometry('400x400')

# Add a light blue canvas drop on the program
canvas = Canvas(mainWin, bg = '#80c1ff', height = 400, width = 400)
canvas.pack()

#Adding frame to the program for aesthetics and formatting
frame = LabelFrame(mainWin,text = 'Please Select From the Following',
                   padx = 30, pady = 30, font = ('Serif', 10, 'bold'), bg = 'gray')
frame.place(relx = 0.5, rely = 0.15, relwidth = 0.575, relheight = 0.65, anchor = 'n')

# Make the functions commands for each button

def openProgram():
    strengthProgram = os.path.join((os.environ.get('HOMEPATH')) , 'Strength_Program')
    programFile = filedialog.askopenfilename(initialdir = strengthProgram, title = 'Select Program.xlsx',
                                     filetypes = (('Excel files', '*.xls*'), ('All Files', '*.*')))
    os.system(programFile)


def generator():
    # Calculations with all the numbers submitted by the user
    currentWeight = weightEntry.get() #This number will be used for protein
    activityLevel = activityEntry.get() #This is for their selected value between 15-17
    percentage = percentageEntry.get() # This is the remaining carbs taken by the percentage

    # Full calculation with activity level of 15
    if activityLevel == '15':
        totalCals = int(currentWeight) * 15 # Formula for total calories
        totalCalsLabel = Label(resultFrame, text = totalCals,
                               bg = '#80c1ff').place(relx = 0.743, rely = 0.097)
        protein = currentWeight ; calsAfterProtein = totalCals - (int(protein) * 4) #Cals after protein
        proteinLabel = Label(resultFrame, text = protein,
                             bg = '#80c1ff').place(relx = 0.75, rely = 0.82)
        if percentage == '60':
            carb = (calsAfterProtein * 0.6) // 4 ; calsAfterCarbs = calsAfterProtein - (calsAfterProtein * 0.6)
            carbLabel = Label(resultFrame, text = carb,
                              bg = '#80c1ff').place(relx = 0.73, rely = 0.598)
            fat = calsAfterCarbs // 9 # Remainder of the calories divided by 9
            fatLabel = Label(resultFrame, text = fat,
                             bg = '#80c1ff').place(relx = 0.75, rely = 0.373)
        elif percentage == '70':
            carb = (calsAfterProtein * 0.7) // 4 ; calsAfterCarbs = calsAfterProtein - (calsAfterProtein * 0.7)
            carbLabel = Label(resultFrame, text = carb,
                              bg = '#80c1ff').place(relx = 0.73, rely = 0.598)
            fat = calsAfterCarbs // 9 # Remainder of the calories divided by 9
            fatLabel = Label(resultFrame, text = fat,
                             bg = '#80c1ff').place(relx = 0.75, rely = 0.373)
        elif percentage == '80':
            carb = (calsAfterProtein * 0.8) // 4 ; calsAfterCarbs = calsAfterProtein - (calsAfterProtein * 0.8)
            carbLabel = Label(resultFrame, text = carb,
                             bg = '#80c1ff').place(relx = 0.73, rely = 0.598)
            fat = calsAfterCarbs // 9 # Remainder of the calories divided by 9
            fatLabel = Label(resultFrame, text = fat,
                             bg = '#80c1ff').place(relx = 0.75, rely = 0.373)

    # Full calculation with activity level of 16
    elif activityLevel == '16':
        totalCals = int(currentWeight) * 16
        totalCalsLabel = Label(resultFrame, text = totalCals,
                               bg = '#80c1ff').place(relx = 0.743, rely = 0.097)
        protein = currentWeight ; calsAfterProtein = totalCals - (int(protein) * 4) #Cals after protein
        proteinLabel = Label(resultFrame, text = protein,
                             bg = '#80c1ff').place(relx = 0.75, rely = 0.82)
        if percentage == '60':
            carb = (calsAfterProtein * 0.6) // 4 ; calsAfterCarbs = calsAfterProtein - (calsAfterProtein * 0.6)
            carbLabel = Label(resultFrame, text = carb,
                              bg = '#80c1ff').place(relx = 0.73, rely = 0.598)
            fat = calsAfterCarbs // 9 # Remainder of the calories divided by 9
            fatLabel = Label(resultFrame, text = fat,
                             bg = '#80c1ff').place(relx = 0.75, rely = 0.373)
        elif percentage == '70':
            carb = (calsAfterProtein * 0.7) // 4 ; calsAfterCarbs = calsAfterProtein - (calsAfterProtein * 0.7)
            carbLabel = Label(resultFrame, text = carb,
                              bg = '#80c1ff').place(relx = 0.73, rely = 0.598)
            fat = calsAfterCarbs // 9 # Remainder of the calories divided by 9
            fatLabel = Label(resultFrame, text = fat,
                             bg = '#80c1ff').place(relx = 0.75, rely = 0.373)
        elif percentage == '80':
            carb = (calsAfterProtein * 0.8) // 4 ; calsAfterCarbs = calsAfterProtein - (calsAfterProtein * 0.8)
            carbLabel = Label(resultFrame, text = carb,
                             bg = '#80c1ff').place(relx = 0.73, rely = 0.598)
            fat = calsAfterCarbs // 9 # Remainder of the calories divided by 9
            fatLabel = Label(resultFrame, text = fat,
                             bg = '#80c1ff').place(relx = 0.75, rely = 0.373)

    # Full calculation of activity level 17
    elif activityLevel == '17':
        totalCals = int(currentWeight) * 17
        totalCalsLabel = Label(resultFrame, text = totalCals,
                               bg = '#80c1ff').place(relx = 0.743, rely = 0.097)
        protein = currentWeight ; calsAfterProtein = totalCals - (int(protein) * 4) #Cals after protein
        proteinLabel = Label(resultFrame, text = protein,
                             bg = '#80c1ff').place(relx = 0.75, rely = 0.82)
        if percentage == '60':
            carb = (calsAfterProtein * 0.6) // 4 ; calsAfterCarbs = calsAfterProtein - (calsAfterProtein * 0.6)
            carbLabel = Label(resultFrame, text = carb,
                              bg = '#80c1ff').place(relx = 0.73, rely = 0.598)
            fat = calsAfterCarbs // 9 # Remainder of the calories divided by 9
            fatLabel = Label(resultFrame, text = fat,
                             bg = '#80c1ff').place(relx = 0.75, rely = 0.373)
        elif percentage == '70':
            carb = (calsAfterProtein * 0.7) // 4 ; calsAfterCarbs = calsAfterProtein - (calsAfterProtein * 0.7)
            carbLabel = Label(resultFrame, text = carb,
                              bg = '#80c1ff').place(relx = 0.73, rely = 0.598)
            fat = calsAfterCarbs // 9 # Remainder of the calories divided by 9
            fatLabel = Label(resultFrame, text = fat,
                             bg = '#80c1ff').place(relx = 0.75, rely = 0.373)
        elif percentage == '80':
            carb = (calsAfterProtein * 0.8) // 4 ; calsAfterCarbs = calsAfterProtein - (calsAfterProtein * 0.8)
            carbLabel = Label(resultFrame, text = carb,
                             bg = '#80c1ff').place(relx = 0.73, rely = 0.598)
            fat = calsAfterCarbs // 9 # Remainder of the calories divided by 9
            fatLabel = Label(resultFrame, text = fat,
                             bg = '#80c1ff').place(relx = 0.75, rely = 0.373)

def newUserPopup():
    newUserSubmit = messagebox.showinfo('Confirm',
                                        'Ok I generated your template, please return to the main menu')
    userProgram = openpyxl.load_workbook('Program - Copy.xlsx')
    sheet = userProgram.active
    programSheet = userProgram['Program']
    sheetFont = Font(bold = True)
    programSheet['M24'].font = sheetFont ; programSheet['M25'].font = sheetFont
    programSheet['M6'].font = sheetFont
    programSheet['M24'] = squatEntry.get()
    programSheet['M25'] = benchEntry.get()
    programSheet['M26'] = deadliftEntry.get()

    userProgram.save('Program - Copy.xlsx')

def newUser():

    newUserWin = Toplevel(mainWin)
    newUserWin.title('Welcome!')
    newUserWin.iconbitmap('barbell.ico')
    newUserWin.geometry('540x400')
    newUserCanvas = Canvas(newUserWin, bg = '#94ffa9', height = 600, width = 600)
    newUserCanvas.place(relx = 0, rely = 0)

    welcomeSign = LabelFrame(newUserWin, text = 'Please Read EVERYTHING Below!'
                        , font = ('Serif', 18, 'bold'),fg = 'blue', labelanchor = 'n', bg = '#94ffa9')
    welcomeSign.grid(row = 0, column = 1, padx = 5)
    # For the user to read and understand the program
    introLabel = Label(welcomeSign, text = """
    Hello and welcome to my strength program, my goal is to hopefully boost your numbers
    at a reasonable pace, avoiding as much strength plateaus as possible. I'm going to coach
    with this strength program that I recently coded, to also benefit myself in trying to master
    Python as a language and for delivering to you, the user, an easily accessible resource that
    you can come back to. Before you can utilize this program, let's first figure out your current
    maxes for your main three lifts. Please fill out the information below, when you are finish, simply
    hit the SUBMIT button.""", bg = '#94ffa9')
    introLabel.pack(pady = 5)
    # Creating the area for them to enter their maxes
    statArea = LabelFrame(newUserWin, text = 'Current Maxes (Squat, Bench, Deadlift)',
                         font = ('Serif', 18, 'bold'),fg = 'blue', labelanchor = 'n', bg = '#94ffa9')
    statArea.grid(row = 5 , column = 1)
    # Label creation for Squat, Bench and Deadlift
    squatLab = Label(statArea, text = 'Squat:', font = ('Serif', 16, 'italic'), bg = '#94ffa9')
    squatLab.grid(row = 1, columnspan = 5, pady = 10)
    benchLab = Label(statArea, text = 'Bench:', font = ('Serif', 16, 'italic'), bg = '#94ffa9')
    benchLab.grid(row = 2 , columnspan = 5, pady = 10)
    deadliftLab = Label(statArea, text = 'Deadlift:', font = ('Serif', 16, 'italic'), bg = '#94ffa9')
    deadliftLab.grid(row = 3, columnspan = 5, pady = 10)

    # Entry creation for Squat, Bench and Deadlift
    global squatEntry
    squatEntry = Entry(statArea)
    squatEntry.place(relx = 0.3, rely = 0.11)
    global benchEntry
    benchEntry = Entry(statArea)
    benchEntry.place(relx = 0.3, rely = 0.43)
    global deadliftEntry
    deadliftEntry = Entry(statArea)
    deadliftEntry.place(relx = 0.3, rely = 0.76)

    # Making a Submit Button for user to press
    submitButton = Button(statArea, text = 'Submit', command = newUserPopup)
    submitButton.place(relx = 0.615 ,rely = 0.1, width = 150, height = 50)

    #Making a button to close current window
    closeButton = Button(statArea, text = 'Cancel', command = newUserWin.destroy)
    closeButton.place(relx = 0.615, rely = 0.55, width = 150, height = 50)



def macronutrients():

    # Create the window when 'My Macronutrients is clicked'
    macroWin = Toplevel(mainWin)
    macroWin.title('My Macronutrients -- Tim\'s Strength Program v 1.0.0 BETA')
    macroWin.iconbitmap('barbell.ico')
    macroWin.geometry('400x450')
    newUserCanvas = Canvas(macroWin, bg = '#80c1ff', height = 600, width = 600)
    newUserCanvas.place(relx = 0, rely = 0)
    # Frames for the macronutrient window
    macroFrame = LabelFrame(macroWin, text = 'Macronutrients Setup', font = ('Serif', 18, 'bold'),
                            labelanchor = 'n', bg = '#80c1ff')
    macroFrame.grid(row = 0, column = 1, padx = 10, pady = 10)
    global resultFrame
    resultFrame = LabelFrame(macroWin, text = 'Your Macronutrients', font = ('Serif', 18, 'bold'),
                             bg = '#80c1ff')
    resultFrame.grid(row = 5, column = 1)
    #Label creation
    macroContent = Label(macroFrame,
                         text = 'Please fill in everything listed below, once done, press Generate',
                         font = ('serif', 10, 'italic'), bg = '#80c1ff')
    macroContent.grid(row = 0, column = 0, pady = 10, columnspan = 3)

    weightQuestion = Label(macroFrame, text = 'Your current weight:', bg = '#80c1ff')
    weightQuestion.grid(row = 1, column = 0, pady = 5, sticky = 'w')
    activityQuestion = Label (macroFrame,
                         text = 'Level of Activity (15(Sedentary)-17(Active)):', bg = '#80c1ff')
    activityQuestion.grid(row = 2, column = 0, pady = 5, sticky = 'w')
    carbsQuestion = Label(macroFrame,
                        text = 'Percentage for Carb intake (60/70/80):', bg = '#80c1ff')
    carbsQuestion.grid(row = 3, column = 0, pady = 5, sticky = 'w')
    lineBreak = Label(macroFrame, text = '', bg = '#80c1ff').grid(row = 4, column = 0, pady = 15)
    calorieLabel = Label(resultFrame, text = 'Calories(kcal):', font = ('serif', 12, 'bold'),
                         bg = '#80c1ff')
    calorieLabel.grid(row = 0, column = 1, pady = 10, sticky = 'w')
    fatLabel = Label(resultFrame, text = 'Fat(g):', font = ('serif', 12, 'bold'), bg = '#80c1ff')
    fatLabel.grid(row =1, column = 1, sticky = 'w')
    carbLabel = Label(resultFrame, text = 'Carbohydrates(g):', font = ('serif', 12, 'bold'),
                      bg = '#80c1ff')
    carbLabel.grid(row = 2, column = 1, sticky = 'w')
    proteinLabel = Label(resultFrame, text = 'Protein(g):', font = ('serirf', 12, 'bold'),
                         bg = '#80c1ff')
    proteinLabel.grid(row = 3, column = 1, sticky = 'w')
    #Button creation for macronutrient window
    generateButton = Button(macroFrame, text = 'Generate', width = '10', command = generator)
    generateButton.place(relx = 0.4, rely = 0.8)
    quitButton = Button(macroWin, text = 'Exit', width = 10
                        , command = macroWin.destroy).place(relx = 0.4, rely = 0.91)
    # Entry creation
    global weightEntry
    global activityEntry
    global percentageEntry
    weightEntry = Entry(macroFrame, width = 10)
    weightEntry.place(relx = 0.76, rely = 0.25)
    activityEntry = Entry(macroFrame, width = 10)
    activityEntry.place(relx = 0.76, rely = 0.425)
    percentageEntry = Entry(macroFrame, width = 10)
    percentageEntry.place(relx = 0.76, rely = 0.6)

# Menu Buttons - New User, My Program, Macronutrients, Quit
newUserBtn = Button(frame, text = 'New User', fg = 'blue', command = newUser)
newUserBtn.place(relx = 0.32, rely = 0)
programBtn = Button(frame, text = 'My Program', fg = 'blue', command = openProgram)
programBtn.place(relx = 0.2, rely = 0.275, relwidth = 0.6)
macroBtn = Button(frame, text = 'My Macronutrients', fg = 'blue', command = macronutrients)
macroBtn.place(relx = 0.04, rely = 0.55, relwidth = 0.9)
exitBtn = Button(frame, text = 'Quit', fg = 'red', command = mainWin.destroy)
exitBtn.place(relx = 0.4, rely = 0.8)




mainWin.mainloop()
