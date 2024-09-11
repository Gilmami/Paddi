from word import *
from ttkwidgets.frames import *
from tkinter import *
from tkinter import ttk
from tkinter import filedialog
from backend import *
from winreg import *


class GUI():

    # Initialize the GUI popup with standard EIS fields
    def __init__(self, *args):
        """Opens the app and populates it with the default values, as well as
        the most commonly selected/used functionality."""
        self.root = Tk()
        self.finished = False
        self.popLoc = 1
        self.function = StringVar(None, "Directory")
        self.proj = StringVar(None, "EIS")
        self.isAgg = BooleanVar(value=False)
        self.breakout = BooleanVar(value=False)
        self.numBreakout = IntVar(self.root, 0)
        self.breakoutDict = {}
        self.numSchoolAgg = IntVar(self.root, 0)
        self.aggDict = {}
        self.country = StringVar(None, "US")
        self.schoolType = StringVar(None, "College")
        self.schoolName = StringVar(None, "School")
        self.schoolName.trace_add("write", self.repop_textboxes)
        self.state = StringVar(None, "State")
        self.root.title("EDU Consulting Project Folder Creation")
        self.root.configure(bg='#f3f0ea', width=600,
                            height=150, cursor="hand2")

        # setting up the area in the window instatiated above for elements to
        # be drawn on, and setting as a grid for lazy auto spacing.
        self.frame = ttk.Frame(self.root, padding=10)
        self.frame.grid()

        # Drawing all the elements with linked commands and placing in grid
        title = Label(self.frame, text="What do you need done?")
        title.grid(column=0, row=self.popLoc)

        jobList = ["Data Only (pdga)", "Directory",
                   "Model", "Templates", "Finals"]
        tipList = ["Pulls in PDGA data survey into your downloads folder.",
                   """Creates directory tree and pulls in timeline, data survey, and utility (EIS only)""", """Imports model from location on Google Drive into your working directory as defined by your inputs.""", """If no directory exists, creates directory and pulls in templates""", """Only pulls final material, directory must exist at this point, or you did something really weird, shame on you."""]
        colCount = 1
        for job in jobList:
            self.jobRadio = Radiobutton(self.frame, text=job,
                                        variable=self.function, value=job)
            self.jobRadio.grid(column=colCount, row=self.popLoc)
            Tooltip(self.jobRadio, text=tipList[colCount-1])
            colCount += 1
        self.popLoc += 1

        LabelProjType = Label(self.frame, text="Select type of project:")
        LabelProjType.grid(column=0, row=self.popLoc)

        projList = ["EIS", "PDGA", "Capital", "VoD", "PSEIS"]
        colCount = 1
        for proj in projList:
            self.projRadio = Radiobutton(self.frame, text=proj,
                                         variable=self.proj, value=proj,
                                         command=self.repop_textboxes)
            self.projRadio.grid(column=colCount, row=self.popLoc)
            colCount += 1
        self.popLoc += 1

        aggText = Label(self.frame, text="Is this an Aggregate?")
        aggText.grid(column=0, row=self.popLoc)
        self.aggRadio = Radiobutton(self.frame, text="Aggregate",
                                    variable=self.isAgg, value=True, command=self.repop_textboxes)
        self.aggRadio.grid(column=1, row=self.popLoc)

        self.indRadio = Radiobutton(
            self.frame, text="Individual", variable=self.isAgg, value=False, command=self.repop_textboxes)
        self.indRadio.grid(column=2, row=self.popLoc)
        self.popLoc += 1

        countryLabel = Label(self.frame, text="Select Country:")
        countryLabel.grid(column=0, row=self.popLoc)

        countryList = ["US", "CAN"]
        colCount = 1
        for country in countryList:
            self.countryRadio = Radiobutton(self.frame, text=country,
                                            variable=self.country,
                                            value=country,
                                            command=self.update_state_dropdown)
            self.countryRadio.grid(column=colCount, row=self.popLoc)
            colCount += 1
        self.popLoc += 1

        self.LabelState = Label(self.frame, text="Select state:")
        self.LabelState.grid(column=0, row=self.popLoc)
        stateList = ["Alabama", "Alaska", "Arizona", "Arkansas",
                     "California", "Colorado", "Connecticut", "Delaware",
                     "District of Columbia", "Florida",
                     "Georgia", "Hawaii", "Idaho", "Illinois", "Indiana", "Iowa", "Kansas",
                     "Kentucky", "Louisiana", "Maine", "Maryland", "Massachusetts",
                     "Michigan", "Minnesota", "Mississippi", "Missouri", "Montana",
                     "Nebraska", "Nevada", "New Hampshire", "New Jersey", "New Mexico",
                     "New York", "North Carolina", "North Dakota", "Ohio", "Oklahoma",
                     "Oregon", "Pennsylvania", "Rhode Island", "South Carolina",
                     "South Dakota", "Tennessee", "Texas", "Utah", "Vermont", "Virginia",
                     "Washington", "West Virginia", "Wisconsin", "Wyoming"]
        self.StateDropdown = ttk.Combobox(
            self.frame, textvariable=self.state, values=stateList)
        self.StateDropdown.grid(column=1, row=self.popLoc, columnspan=2)
        self.state.trace_add("write", self.update_school_dropdown)
        self.popLoc += 1

        self.SchoolTypeLabel = Label(self.frame, text="Type of institution:")
        self.SchoolTypeLabel.grid(column=0, row=self.popLoc)

        typeList = ["University-Public", "University-Private", "College"]
        colCount = 1
        for entry in typeList:
            self.typeRadio = Radiobutton(self.frame, text=entry,
                                         variable=self.schoolType, value=entry)
            self.typeRadio.grid(column=colCount, row=self.popLoc)
            colCount += 1
        self.popLoc += 1

        breakoutText = Label(self.frame, text="Any campus breakouts?")
        breakoutText.grid(column=0, row=self.popLoc)
        self.breakoutRadio = Radiobutton(self.frame, text="Yes", variable=self.breakout, value=True,
                                         command=self.repop_textboxes)
        self.breakoutRadio.grid(column=1, row=self.popLoc)
        self.noBreakoutRadio = Radiobutton(self.frame, text="No", variable=self.breakout, value=False,
                                           command=self.repop_textboxes)
        self.noBreakoutRadio.grid(column=2, row=self.popLoc)
        self.popLoc += 1

        self.schoolLabel = Label(self.frame, text="Select a school:")
        self.schoolLabel.grid(column=0, row=self.popLoc)

        schoolList = get_schools(self.country.get(), self.state.get())
        self.schoolDropdown = ttk.Combobox(
            self.frame, textvariable=self.schoolName,
            values=schoolList.sort())
        self.schoolDropdown.grid(column=1, row=self.popLoc, columnspan=2)
        self.popLoc += 1

        self.schoolAcronymLabel = Label(
            self.frame, text="Input School Acronym:")
        self.schoolAcronymLabel.grid(column=0, row=self.popLoc)

        self.schoolAcronymText = Text(self.frame, height=1, width=20)
        self.schoolAcronymText.bind("<Tab>", self.select_next)
        self.schoolAcronymText.bind("<Shift-KeyPress-Tab>", self.select_prev)
        self.schoolAcronymText.grid(column=1, row=self.popLoc, columnspan=2)
        self.popLoc += 1

        self.AnalysisYearLabel = Label(self.frame, text="Input Analysis Year:")
        self.AnalysisYearLabel.grid(column=0, row=self.popLoc)
        self.AnalysisYearText = Text(self.frame, height=1, width=20)
        self.AnalysisYearText.bind("<Tab>", self.select_next)
        self.AnalysisYearText.bind("<Shift-KeyPress-Tab>", self.select_prev)
        self.AnalysisYearText.grid(column=1, row=self.popLoc, columnspan=2)

        self.popLoc += 1
        self.runButton = Button(self.frame, text="Run", command=self.run)
        self.runButton.grid(column=3, row=self.popLoc)
        self.root.mainloop()

    def update_state_dropdown(self, *args):
        """Updates the state dropdown and changes to province if Canada is
        select_nexted as country."""
        self.popLoc = 5
        for w in self.frame.grid_slaves(5):
            w.grid_forget()
        if self.country.get() == "CAN":
            self.LabelProv = Label(self.frame, text="Select Province:")
            self.LabelProv.grid(column=0, row=self.popLoc)

            provinceList = ["Alberta", "British Columbia", "Manitoba",
                            "New Brunswick", "Newfoundland and Labrador",
                            "Northwest Territories", "Nova Scotia", "Nunavut", "Ontario",
                            "Prince Edward Island", "Quebec", "Saskatchewan", "Yukon"]

            self.state = StringVar()
            self.state.set("Province")
            self.ProvDropdown = ttk.Combobox(
                self.frame, textvariable=self.state, values=provinceList)
            self.state.trace_add("write", self.update_school_dropdown)
            self.ProvDropdown.grid(column=1, row=self.popLoc, columnspan=2)

        else:
            self.LabelState = Label(self.frame, text="Select state:")
            self.LabelState.grid(column=0, row=self.popLoc)

            stateList = ["Alabama", "Alaska", "Arizona", "Arkansas",
                         "California", "Colorado", "Connecticut", "Delaware",
                         "District of Columbia", "Florida",
                         "Georgia", "Hawaii", "Idaho", "Illinois", "Indiana", "Iowa", "Kansas",
                         "Kentucky", "Louisiana", "Maine", "Maryland", "Massachusetts",
                         "Michigan", "Minnesota", "Mississippi", "Missouri", "Montana",
                         "Nebraska", "Nevada", "New Hampshire", "New Jersey", "New Mexico",
                         "New York", "North Carolina", "North Dakota", "Ohio", "Oklahoma",
                         "Oregon", "Pennsylvania", "Rhode Island", "South Carolina",
                         "South Dakota", "Tennessee", "Texas", "Utah", "Vermont", "Virginia",
                         "Washington", "West Virginia", "Wisconsin", "Wyoming"]

            self.state = StringVar()
            self.state.set("State")
            self.StateDropdown = ttk.Combobox(
                self.frame, textvariable=self.state, values=stateList)
            self.state.trace_add("write", self.update_school_dropdown)

            self.StateDropdown.grid(column=1, row=self.popLoc, columnspan=2)

    def select_next(self, event):
        """Helper function to select next widget, allows for better navigation
        in the app."""
        event.widget.tk_focusNext().focus()
        return("break")

    def select_prev(self, event):
        """Helper function to select previous widget, allows for better navigation
        in the app."""
        event.widget.tk_focusPrev().focus()
        return("break")

    def update_school_dropdown(self, state, something, method):
        """Updates the school dropdown using teh state and looks at the
        filepath using get_schools method to populate based on existing
        directories."""
        if self.isAgg.get():
            pass
        else:
            # print("update_school_dropdown")
            self.popLoc = 8
            for w in self.frame.grid_slaves(row=self.popLoc):
                w.grid_forget()
            schoolList = get_schools(self.country.get(), self.state.get())
            self.schoolLabel = Label(self.frame, text="Select School:")
            self.schoolLabel.grid(column=0, row=self.popLoc)
            self.schoolName = StringVar(None, "School")
            self.schoolName.trace_add("write", self.repop_textboxes)
            self.schoolDropdown = ttk.Combobox(
                self.frame, textvariable=self.schoolName, values=schoolList)
            self.schoolDropdown.grid(column=1, row=self.popLoc, columnspan=2)

    def retrieve_textbox(self, textbox):
        """this just makes it easier to get teh analysis year."""
        input = textbox.get("1.0", 'end-1c')
        return(input)

    def grid_forget(self, int1, int2):
        """Clears screen widgets between int1 and int2 to repopulate with other
        widgets."""
        for i in range(int1, int2):
            for w in self.frame.grid_slaves(i):
                w.grid_forget()

    def box_gen(self, box_number, campus_or_school):
        """Generates textboxes and dynamically assigns them a variable for
        aggregates based on the number of schools the user selects."""
        self.boxDict = {}
        try:
            for num in range(0, box_number):
                self.school = Label(
                    self.frame, text=f"Input {campus_or_school} name:")
                self.school.grid(column=0, row=self.popLoc)

                self.boxDict[campus_or_school +
                             str(num)] = Text(self.frame, height=1, width=20)
                self.boxDict[campus_or_school +
                             str(num)].bind("<Tab>", self.select_next)
                self.boxDict[campus_or_school +
                             str(num)].bind("<Shift-KeyPress-Tab>", self.select_prev)
                self.boxDict[campus_or_school + str(num)].grid(
                    column=1, row=self.popLoc, columnspan=2)
                self.popLoc += 1
        except:
            pass
        return(self.boxDict)

    def repop_textboxes(self, *args):
        """Dynamically repopulates the screen based on user selection in a
        number of locations, resulting in the app feeling like it's actually
        coded by someone who actually knows what they're doing."""

        self.popLoc = 8
        if self.isAgg.get() == True and self.breakout.get() == False:
            print("agg yes, breakout no")
            try:
                self.grid_forget(self.popLoc, 100)
                self.labelAggNum = Label(
                    self.frame, text="How many Schools in Agg?")
                self.labelAggNum .grid(column=0, row=self.popLoc)
                aggNumList = list()
                for i in range(1, 11):
                    aggNumList.append(i)
                self.aggDropdown = ttk.Combobox(
                    self.frame, textvariable=self.numSchoolAgg, values=aggNumList)
                self.aggDropdown.grid(column=1, row=self.popLoc, columnspan=2)
                self.numSchoolAgg.trace_add("write", self.repop_textboxes)
                self.popLoc += 1

                self.aggName = Label(
                    self.frame, text="Input Aggregate Name:")
                self.aggName.grid(column=0, row=self.popLoc)

                self.aggText = Text(self.frame, height=1, width=20)
                self.aggText.bind("<Tab>", self.select_next)
                self.aggText.bind("<Shift-KeyPress-Tab>", self.select_prev)
                self.aggText.grid(
                    column=1, row=self.popLoc, columnspan=2)
                self.popLoc += 1

                self.AnalysisYearLabel = Label(
                    self.frame, text="Input Analysis Year:")
                self.AnalysisYearLabel.grid(column=0, row=self.popLoc)

                self.AnalysisYearText = Text(self.frame, height=1, width=20)
                self.AnalysisYearText.bind("<Tab>", self.select_next)
                self.AnalysisYearText.bind(
                    "<Shift-KeyPress-Tab>", self.select_prev)
                self.AnalysisYearText.grid(
                    column=1, row=self.popLoc, columnspan=2)
                self.popLoc += 1

                self.aggDict = self.box_gen(self.numSchoolAgg.get(), "school")
                print(self.aggDict)

                self.runButton = Button(
                    self.frame, text="Run", command=self.run)
                self.runButton.grid(column=3, row=self.popLoc)
            except:
                pass
        elif self.isAgg.get() == False and self.breakout.get() == True:
            print("breakout only")
            self.popLoc = 10
            self.grid_forget(self.popLoc, 100)
            try:
                self.labelBreakoutNum = Label(
                    self.frame, text="How many breakouts?")
                self.labelBreakoutNum .grid(column=0, row=self.popLoc)
                breakoutList = list()
                for i in range(1, 11):
                    breakoutList.append(i)
                self.breakoutDropdown = ttk.Combobox(
                    self.frame, textvariable=self.numBreakout,
                    values=breakoutList)
                self.breakoutDropdown.grid(
                    column=1, row=self.popLoc, columnspan=2)
                self.numBreakout.trace_add("write", self.repop_textboxes)
                self.popLoc += 1

                self.breakoutDict = self.box_gen(
                    self.numBreakout.get(), "campus")

                self.AnalysisYearLabel = Label(
                    self.frame, text="Input Analysis Year:")
                self.AnalysisYearLabel.grid(column=0, row=self.popLoc)

                self.AnalysisYearText = Text(self.frame, height=1, width=20)
                self.AnalysisYearText.bind("<Tab>", self.select_next)
                self.AnalysisYearText.bind(
                    "<Shift-KeyPress-Tab>", self.select_prev)
                self.AnalysisYearText.grid(
                    column=1, row=self.popLoc, columnspan=2)
                self.popLoc += 1

                self.runButton = Button(
                    self.frame, text="Run", command=self.run)
                self.runButton.grid(column=3, row=self.popLoc)
                print(self.breakoutDict)

            except:
                pass

        elif self.isAgg.get() == True and self.breakout.get() == True:
            print("agg and breakout")
            try:
                self.grid_forget(self.popLoc, 100)
                self.labelAggNum = Label(
                    self.frame, text="How many Schools in Agg?")
                self.labelAggNum .grid(column=0, row=self.popLoc)
                aggNumList = list()
                for i in range(1, 11):
                    aggNumList.append(i)
                self.aggDropdown = ttk.Combobox(
                    self.frame, textvariable=self.numSchoolAgg, values=aggNumList)
                self.aggDropdown.grid(column=1, row=self.popLoc, columnspan=2)
                self.numSchoolAgg.trace_add("write", self.repop_textboxes)
                self.popLoc += 1

                self.aggName = Label(
                    self.frame, text="Input Aggregate Name:")
                self.aggName.grid(column=0, row=self.popLoc)

                self.aggText = Text(self.frame, height=1, width=20)
                self.aggText.bind("<Tab>", self.select_next)
                self.aggText.bind("<Shift-KeyPress-Tab>", self.select_prev)
                self.aggText.grid(
                    column=1, row=self.popLoc, columnspan=2)
                self.popLoc += 1

                self.AnalysisYearLabel = Label(
                    self.frame, text="Input Analysis Year:")
                self.AnalysisYearLabel.grid(column=0, row=self.popLoc)

                self.AnalysisYearText = Text(self.frame, height=1, width=20)
                self.AnalysisYearText.bind("<Tab>", self.select_next)
                self.AnalysisYearText.bind(
                    "<Shift-KeyPress-Tab>", self.select_prev)
                self.AnalysisYearText.grid(
                    column=1, row=self.popLoc, columnspan=2)
                self.popLoc += 1

                self.aggDict = self.box_gen(self.numSchoolAgg.get(), "school")

                self.labelBreakoutNum = Label(
                    self.frame, text="How many breakouts?")
                self.labelBreakoutNum .grid(column=0, row=self.popLoc)
                breakoutList = list()
                for i in range(1, 11):
                    breakoutList.append(i)
                self.breakoutDropdown = ttk.Combobox(
                    self.frame, textvariable=self.numBreakout,
                    values=breakoutList)
                self.breakoutDropdown.grid(
                    column=1, row=self.popLoc, columnspan=2)
                self.numBreakout.trace_add("write", self.repop_textboxes)
                self.popLoc += 1

                self.breakoutDict = self.box_gen(
                    self.numBreakout.get(), "campus")
                print(self.breakoutDict)
                print(self.aggDict)

                self.runButton = Button(
                    self.frame, text="Run", command=self.run)
                self.runButton.grid(column=3, row=self.popLoc)

            except:
                pass
        else:
            print("final else")
            print(self.popLoc)
            self.grid_forget(self.popLoc, 100)
            self.schoolLabel = Label(self.frame, text="Select a school:")
            self.schoolLabel.grid(column=0, row=self.popLoc)

            schoolList = get_schools(self.country.get(), self.state.get())
            self.schoolDropdown = ttk.Combobox(
                self.frame, textvariable=self.schoolName, values=schoolList)
            self.schoolName.trace_add("write", self.repop_textboxes)
            self.schoolDropdown.grid(column=1, row=self.popLoc, columnspan=2)
            self.popLoc += 1

        if self.proj.get() in ["EIS", "Capital", "VoD", "PSEIS"] and self.isAgg.get() == False and self.breakout.get() == False:
            self.grid_forget(self.popLoc, 100)
            self.schoolAcronymLabel = Label(
                self.frame, text="Input School Acronym:")
            self.schoolAcronymLabel.grid(column=0, row=self.popLoc)

            self.schoolAcronymText = Text(self.frame, height=1, width=20)
            self.schoolAcronymText.bind("<Tab>", self.select_next)
            self.schoolAcronymText.bind(
                "<Shift-KeyPress-Tab>", self.select_prev)
            self.schoolAcronymText.grid(
                column=1, row=self.popLoc, columnspan=2)
            self.popLoc += 1

            self.AnalysisYearLabel = Label(
                self.frame, text="Input Analysis Year:")
            self.AnalysisYearLabel.grid(column=0, row=self.popLoc)

            self.AnalysisYearText = Text(self.frame, height=1, width=20)
            self.AnalysisYearText.bind("<Tab>", self.select_next)
            self.AnalysisYearText.bind(
                "<Shift-KeyPress-Tab>", self.select_prev)
            self.AnalysisYearText.grid(
                column=1, row=self.popLoc, columnspan=2)
            self.popLoc += 1

            self.runButton = Button(self.frame, text="Run", command=self.run)
            self.runButton.grid(column=3, row=self.popLoc)

        elif self.proj.get() == "PDGA" and self.isAgg.get() == False and self.breakout.get() == False:
            self.grid_forget(self.popLoc, 100)
            self.schoolAcronymLabel = Label(
                self.frame, text="Input School Acronym:")
            self.schoolAcronymLabel.grid(column=0, row=self.popLoc)
            self.schoolAcronymText = Text(self.frame, height=1, width=20)
            self.schoolAcronymText.bind("<Tab>", self.select_next)
            self.schoolAcronymText.bind(
                "<Shift-KeyPress-Tab>", self.select_prev)
            self.schoolAcronymText.grid(
                column=1, row=self.popLoc, columnspan=2)
            self.popLoc += 1

            self.AnalysisYearLabel = Label(
                self.frame, text="Input Analysis Year:")
            self.AnalysisYearLabel.grid(column=0, row=self.popLoc)
            self.AnalysisYearText = Text(self.frame, height=1, width=20)
            self.AnalysisYearText.bind("<Tab>", self.select_next)
            self.AnalysisYearText.bind(
                "<Shift-KeyPress-Tab>", self.select_prev)
            self.AnalysisYearText.grid(
                column=1, row=self.popLoc, columnspan=2)
            self.popLoc += 1

            self.IDLabel = Label(self.frame, text="Institution ID")
            self.IDLabel.grid(column=0, row=self.popLoc)
            self.ID = Text(self.frame, height=1, width=20)
            self.ID.bind("<Tab>", self.select_next)
            self.ID.bind("<Shift-KeyPress-Tab>", self.select_prev)
            self.ID.grid(column=1, row=self.popLoc, columnspan=2)
            self.popLoc += 1

            self.dataRunLabel = Label(self.frame, text="Data Run")
            self.dataRunLabel.grid(column=0, row=self.popLoc)
            self.dataRun = Text(self.frame, height=1, width=20)
            self.dataRun.bind("<Tab>", self.select_next)
            self.dataRun.bind("<Shift-KeyPress-Tab>", self.select_prev)
            self.dataRun.grid(column=1, row=self.popLoc, columnspan=2)
            self.popLoc += 1

            self.completionsLabel = Label(
                self.frame, text="Completions Base Year")
            self.completionsLabel.grid(column=0, row=self.popLoc)
            self.completionsLabel.grid(column=0, row=self.popLoc)
            self.completions = Text(self.frame, height=1, width=20)
            self.completions.bind("<Tab>", self.select_next)
            self.completions.bind("<Shift-KeyPress-Tab>", self.select_prev)
            self.completions.grid(
                column=1, row=self.popLoc, columnspan=2)
            self.popLoc += 1

            self.runButton = Button(
                self.frame, text="Run", command=self.run)
            self.runButton.grid(column=3, row=self.popLoc)

    def school_path(self):
        if self.function.get() == "Directory":
            path = os.path.normpath("G:\\Shared Drives\\EDU Consulting - Private\\Completed Reports"
                                    + "\\" + self.country.get() + "\\" +
                                    self.state.get() + "\\" +
                                    self.schoolName.get() + "\\" +
                                    self.proj.get() + "\\" +
                                    self.retrieve_textbox(self.AnalysisYearText))
        elif self.function.get() == "Templates":
            path = os.path.normpath("G:\\Shared Drives\\EDU Consulting - Private\\Completed Reports"
                                    + "\\" + self.country.get() + "\\" +
                                    self.state.get() + "\\" +
                                    self.schoolName.get() + "\\" +
                                    self.proj.get() + "\\" +
                                    self.retrieve_textbox(self.AnalysisYearText) +
                                    "\\" + "Drafts")

        elif self.function.get() == "Finals":
            path = os.path.normpath("G:\\Shared Drives\\EDU Consulting - Private\\Completed Reports"
                                    + "\\" + self.country.get() + "\\" +
                                    self.state.get() + "\\" +
                                    self.schoolName.get() + "\\" +
                                    self.proj.get() + "\\" +
                                    self.retrieve_textbox(self.AnalysisYearText) +
                                    "\\" + "Finals")

        elif self.function.get() == "Data Only (pdga)":
            path = os.path.normpath(os.getenv("USERPROFILE") + "\\Downloads")

        else:
            path = None

        return(path)

    def kill(self):
        """this one is for the lulz."""
        path = self.school_path()
        self.root.destroy()
        try:
            os.startfile(path)
        except:
            pass

    def make_list(self, aggDict):
        """Helper function to convert a dictionary constructed within the UI
        into a list of names to be passed to another function."""
        print(aggDict)
        aggList = list()
        for name in aggDict:
            aggList.append(self.retrieve_textbox(aggDict[name]))
        return(aggList)

    def run(self):
        """This runs different backend functions and passes user inputs to
        those functions and returns a message when the process has completed or
        something has broken, if all goes well."""
        self.finished = False

        if self.isAgg.get() == False and self.breakout.get() == False:
            if self.function.get() == "Directory" and len(self.retrieve_textbox(self.AnalysisYearText)) == 4 and len(self.retrieve_textbox(self.schoolAcronymText)) != 0:
                make_proj_tree(self.country.get(), self.state.get(),
                               self.schoolName.get(),
                               self.retrieve_textbox(self.AnalysisYearText), self.proj.get())
                try:
                    import_admin(self.country.get(), self.state.get(),
                                 self.schoolName.get(),
                                 self.retrieve_textbox(
                                     self.AnalysisYearText),
                                 self.retrieve_textbox(self.schoolAcronymText),
                                 self.proj.get(), self.schoolType.get(),
                                 self.retrieve_textbox(self.ID),
                                 self.retrieve_textbox(self.dataRun),
                                 self.retrieve_textbox(self.completions))
                except:
                    import_admin(self.country.get(), self.state.get(),
                                 self.schoolName.get(),
                                 self.retrieve_textbox(
                                     self.AnalysisYearText),
                                 self.retrieve_textbox(self.schoolAcronymText),
                                 self.proj.get(), self.schoolType.get())
                    print(self.retrieve_textbox(self.AnalysisYearText))
                self.finished = True
                self.messageText = "Successfully completed Directory creation and admin import for " + \
                    self.schoolName.get() + " in analysis year " + \
                    self.retrieve_textbox(self.AnalysisYearText) + "!"

            elif self.function.get() == "Templates" and len(self.retrieve_textbox(self.AnalysisYearText)) == 4 and len(self.retrieve_textbox(self.schoolAcronymText)) != 0:
                pathList = import_templates(self.proj.get(), self.schoolType.get(),
                                            self.country.get(),
                                            self.state.get(), self.schoolName.get(),
                                            self.retrieve_textbox(
                                                self.AnalysisYearText),
                                            self.retrieve_textbox(self.schoolAcronymText))
                modelPathDict = find_model(self.country.get(),
                                           self.state.get(),
                                           [self.schoolName.get()],
                                           self.proj.get(),
                                           self.retrieve_textbox(
                    self.AnalysisYearText))
                for school in modelPathDict:
                    for report in pathList:
                        change_source(modelPathDict[school][0],
                                      modelPathDict[school][1], pathList[report])

                print(self.retrieve_textbox(self.schoolAcronymText))
                self.finished = True
                self.messageText = "Successfully completed template import for " + \
                    self.schoolName.get() + " in analsyis year " + \
                    self.retrieve_textbox(self.AnalysisYearText) + "!"

            elif self.function.get() == "Finals" and len(self.retrieve_textbox(self.AnalysisYearText)) == 4 and len(self.retrieve_textbox(self.schoolAcronymText)) != 0:
                pathList = import_finals(self.proj.get(), self.schoolType.get(),
                                         self.country.get(),
                                         self.state.get(), self.schoolName.get(),
                                         self.retrieve_textbox(
                                             self.AnalysisYearText),
                                         self.retrieve_textbox(self.schoolAcronymText))
                modelPathDict = find_model(self.country.get(),
                                           self.state.get(),
                                           [self.schoolName.get()],
                                           self.proj.get(),
                                           self.retrieve_textbox(
                    self.AnalysisYearText))
                for school in modelPathDict:
                    for report in pathList:
                        change_source(modelPathDict[school][0],
                                      modelPathDict[school][1], pathList[report])
                print(self.retrieve_textbox(self.schoolAcronymText))
                self.finished = True
                self.messageText = "Successfully completed finals import for " + \
                    self.schoolName.get() + " in analysis year " + \
                    self.retrieve_textbox(self.AnalysisYearText) + "!"

            elif self.function.get() == "Data Only (pdga)":
                with OpenKey(HKEY_CURRENT_USER, 'SOFTWARE\Microsoft\Windows\CurrentVersion\Explorer\Shell Folders') as key:
                    Downloads = QueryValueEx(
                        key, '{374DE290-123F-4565-9164-39C4925E467B}')[0]
                path = Downloads + "\\" + self.retrieve_textbox(
                    self.schoolAcronymText) + "_PDGA_DataSurvey_" + self.retrieve_textbox(self.completions) + ".xlsx"
                get_pdga_data_survey(self.retrieve_textbox(self.ID),
                                     self.retrieve_textbox(self.dataRun),
                                     self.retrieve_textbox(self.completions),
                                     path)
                self.finished = True
                self.messageText = "Successfully downloaded data survey for IPEDS ID: " + self.retrieve_textbox(
                    self.ID) + " for data run: " + self.retrieve_textbox(self.dataRun) + ". Your file can be found in your downloads folder."

            elif self.function.get() == "Model":
                import_model(self.country.get(), self.state.get(),
                             [self.schoolName.get()],
                             self.proj.get(),
                             self.retrieve_textbox(self.schoolAcronymText),
                             self.retrieve_textbox(self.AnalysisYearText))
                self.messageText = f"Successfully imported model for {self.schoolName.get()}!"
                self.finished = True

            elif len(self.retrieve_textbox(self.AnalysisYearText)) != 4 and len(self.retrieve_textbox(self.schoolAcronymText)) != 0:
                print("analysis Year incorrect format (needs to be 4 numbers)")
                pass

            elif len(self.retrieve_textbox(self.AnalysisYearText)) == 4 and len(self.retrieve_textbox(self.schoolAcronymText)) == 0:
                print("No school acronym inputted, Fix it and try again.")
                pass

            elif len(self.retrieve_textbox(self.AnalysisYearText)) != 4 and len(self.retrieve_textbox(self.schoolAcronymText)) == 0:
                print(
                    "Now you're just being silly. Wrong analysis year format and no acronym!?")
                pass

            else:
                print("Some other error occured in non agg, existing school loop")
                pass

        elif self.isAgg.get() == True and self.breakout.get() == False:
            if self.function.get() == "Directory" and len(self.retrieve_textbox(self.AnalysisYearText)) == 4 and len(self.retrieve_textbox(self.aggText)) != 0:
                make_agg_tree(self.country.get(), self.state.get(),
                              self.retrieve_textbox(self.aggText),
                              self.make_list(self.aggDict),
                              self.retrieve_textbox(self.AnalysisYearText),
                              self.proj.get())
                self.messageText = "Aggregate directory function run for " + \
                    str(self.numSchoolAgg.get()) + " schools."

            elif self.function.get() == "Templates" and len(self.retrieve_textbox(self.AnalysisYearText)) == 4 and len(self.retrieve_textbox(self.aggText)) != 0:
                pathList = import_templates(self.proj.get(), self.schoolType.get(),
                                            self.country.get(), self.state.get(),
                                            self.make_list(
                                                self.aggDict),
                                            self.retrieve_textbox(
                                                self.AnalysisYearText),
                                            "placeholder",
                                            self.retrieve_textbox(self.aggText))
                modelPathDict = find_model(self.country.get(), self.state.get(),
                                           self.make_list(
                    self.aggDict),
                    self.proj.get(), self.retrieve_textbox(
                    self.AnalysisYearText), self.aggName.get())
                for school in modelPathDict:
                    for report in pathList:
                        change_source(modelPathDict[school][0],
                                      modelPathDict[school][1], pathList[report])

                self.messageText = "Aggregate Templates run for " + \
                    str(self.numSchoolAgg.get()) + " schools."

            elif self.function.get() == "Finals" and len(self.retrieve_textbox(self.AnalysisYearText)) == 4 and len(self.retrieve_textbox(self.aggText)) != 0:
                pathList = import_finals(self.proj.get(), self.schoolType.get(),
                                         self.country.get(),
                                         self.state.get(),
                                         self.make_list(
                                             self.aggDict),
                                         self.retrieve_textbox(
                                             self.AnalysisYearText),
                                         "placeholder", self.retrieve_textbox(self.aggText))

                modelPathDict = find_model(self.country.get(), self.state.get(),
                                           self.make_list(
                    self.aggDict),
                    self.proj.get(), self.retrieve_textbox(
                    self.AnalysisYearText), self.aggName.get())
                for school in modelPathDict:
                    for report in pathList:
                        change_source(modelPathDict[school][0],
                                      modelPathDict[school][1], pathList[report])

            elif self.function.get() == "Data Only (pdga)":
                self.messageText = "What are you even thinking? Goofer."

            elif self.function.get() == "Model":
                import_model(self.country.get(), self.state.get(),
                             self.make_list(self.aggDict),
                             self.proj.get(),
                             year=self.retrieve_textbox(self.AnalysisYearText),
                             aggName=self.aggName.get())
                self.messageText = f"Model successfully imported for {self.aggName.get()}."
            else:
                self.messageText = "Something's wrong. Figure it out smarty pants."

        elif self.isAgg.get() == False and self.breakout.get() == True:
            if self.function.get() == "Directory" and len(self.retrieve_textbox(self.AnalysisYearText)) == 4 and len(self.retrieve_textbox(self.schoolAcronymText)) != 0:
                make_proj_tree(self.country.get(), self.state.get(),
                               self.schoolName.get(),
                               self.retrieve_textbox(self.AnalysisYearText),
                               self.proj.get(),
                               campuses=self.make_list(self.breakoutDict))
                try:
                    import_admin(self.country.get(), self.state.get(),
                                 self.schoolName.get(),
                                 self.retrieve_textbox(
                                     self.AnalysisYearText),
                                 self.retrieve_textbox(self.schoolAcronymText),
                                 self.proj.get(), self.schoolType.get(),
                                 self.retrieve_textbox(self.ID),
                                 self.retrieve_textbox(self.dataRun),
                                 self.retrieve_textbox(self.completions), campuses=self.make_list(self.breakoutDict))
                except:
                    import_admin(self.country.get(), self.state.get(),
                                 self.schoolName.get(),
                                 self.retrieve_textbox(
                                     self.AnalysisYearText),
                                 self.retrieve_textbox(self.schoolAcronymText),
                                 self.proj.get(), self.schoolType.get(), campuses=self.make_list(self.breakoutDict))
                    print(self.retrieve_textbox(self.AnalysisYearText))
                self.finished = True
                self.messageText = "Successfully completed Directory creation and admin import for " + \
                    self.schoolName.get() + " in analysis year " + \
                    self.retrieve_textbox(self.AnalysisYearText) + "!"

            elif self.function.get() == "Templates" and len(self.retrieve_textbox(self.AnalysisYearText)) == 4 and len(self.retrieve_textbox(self.aggText)) != 0:
                pathList = import_templates(self.proj.get(), self.schoolType.get(),
                                            self.country.get(), self.state.get(),
                                            self.schoolName.get(),
                                            self.retrieve_textbox(
                                                self.AnalysisYearText),
                                            "placeholder",
                                            campuses=self.make_list(self.breakoutDict))
                modelPathDict = find_model(self.country.get(), self.state.get(),
                                           self.schoolName.get(),
                                           self.proj.get(), self.retrieve_textbox(
                    self.AnalysisYearText), campuses=self.make_list(self.breakoutDict))
                for school in modelPathDict:
                    for report in pathList:
                        change_source(modelPathDict[school][0],
                                      modelPathDict[school][1], pathList[report])

                self.messageText = "Templates import and source change for\
                campuses of " + self.schoolName.get()

            elif self.function.get() == "Finals" and len(self.retrieve_textbox(self.AnalysisYearText)) == 4 and len(self.retrieve_textbox(self.aggText)) != 0:
                pathList = import_finals(self.proj.get(), self.schoolType.get(),
                                         self.country.get(),
                                         self.state.get(),
                                         self.schoolName.get(),
                                         self.retrieve_textbox(
                                             self.AnalysisYearText),
                                         "placeholder",
                                         campuses=self.make_list(self.breakoutDict))

                modelPathDict = find_model(self.country.get(), self.state.get(),
                                           self.schoolName.get(),
                                           self.proj.get(), self.retrieve_textbox(
                    self.AnalysisYearText), campuses=self.make_list(self.breakoutDict))
                for school in modelPathDict:
                    for report in pathList:
                        change_source(modelPathDict[school][0],
                                      modelPathDict[school][1], pathList[report])

            elif self.function.get() == "Data Only (pdga)":
                self.messageText = "What are you even thinking? Goofer."

            elif self.function.get() == "Model":
                import_model(self.country.get(), self.state.get(),
                             self.schoolName.get(),
                             self.proj.get(),
                             year=self.retrieve_textbox(self.AnalysisYearText),
                             campuses=self.make_list(self.breakoutDict))
                self.messageText = f"Model successfully imported for\
                        {self.schoolName.get()}'s campuses."
            else:
                self.messageText = "Something's wrong. Figure it out smarty pants."
        else:
            self.messageText = "Something unknown went wrong outside of new, \
                    existing, breakout, and agg functions. Light the beacons."

        print(self.finished)
        if self.finished == True:
            self.grid_forget(1, 100)
            self.popLoc = 1
            self.message = Label(self.frame, text=self.messageText)
            self.message.grid(row=self.popLoc, column=2)
            self.popLoc += 1

            self.okButton = Button(
                self.frame, text="Go away.", command=self.kill)
            self.okButton.grid(row=self.popLoc, column=2)
        else:
            self.finishMessage = Label(
                self.frame, text=self.messageText)
            self.finishMessage.grid(row=self.popLoc, column=0, columnspan=5)


if __name__ == "__main__":
    GUI()
    quit()
