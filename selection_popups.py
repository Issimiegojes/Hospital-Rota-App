
from tkinter import *

def prefer_count(prefer_popup_inputs):
    
    root = prefer_popup_inputs["give_root"]
    days_list = prefer_popup_inputs["give_days_list"]
    starting_weekday = prefer_popup_inputs["give_starting_weekday"]
    worker_rows = prefer_popup_inputs["give_worker_rows"]
    row_num = prefer_popup_inputs["give_row_num"]
    selected_prefer_days = prefer_popup_inputs["give_selected_prefer_days"]
    error_label = prefer_popup_inputs["give_error_label"]

    popup = Toplevel(root)
    popup.title("Select Days Prefer Work")

    # Create a canvas and scrollbar for scrolling
    canvas = Canvas(popup)
    scrollbar = Scrollbar(popup, orient="vertical", command=canvas.yview)
    scrollable_frame = Frame(canvas)

    scrollable_frame.bind(  # .bind(event,function) This is needed, because when popup appears, it is empty, and only later gets populated by checklist
        "<Configure>", # Event of a widget changing its size
        lambda e: canvas.configure(scrollregion=canvas.bbox("all")) # canvas.configure(scrollregion= means "this is your scroll area. bbox=BindingBox, calculates total dimensions of everything inside canvas. 
    )
    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw") # The Canvas normally displays drawings — shapes, lines, images. To put a regular tkinter widget (like a Frame) inside a canvas, you use create_window() — it embeds the widget as if it were a drawn object.
    canvas.configure(yscrollcommand=scrollbar.set)  # Canvas tells scrollbar where it is, if scrolled with mousewheel for example

    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    def on_popup_mousewheel(event):
        canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    popup.bind("<MouseWheel>", on_popup_mousewheel)
    canvas.bind("<MouseWheel>", on_popup_mousewheel)

    Label(scrollable_frame, text="Day").grid(row=0, column=0)
    Label(scrollable_frame, text="Weekday").grid(row=0, column=1)
    Label(scrollable_frame, text="Day Shift").grid(row=0, column=2)
    Label(scrollable_frame, text="Night Shift").grid(row=0, column=3)

    day_names = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
    check_vars = []

    # Create separate lists to hold ONLY day shift variables and ONLY night shift variables
    day_shift_vars = []  # Empty basket for day shift checkboxes
    night_shift_vars = []  # Empty basket for night shift checkboxes

    for i, day in enumerate(days_list, start=1):
        weekday = (starting_weekday + (day - 1)) % 7
        day_name = day_names[weekday]
        Label(scrollable_frame, text=str(day)).grid(row=i, column=0)
        Label(scrollable_frame, text=day_name).grid(row=i, column=1)
        
        # Day shift checkbox
        day_var = IntVar()
        Checkbutton(scrollable_frame, variable=day_var).grid(row=i, column=2)
        check_vars.append((f"Day {day}", day_var))
        day_shift_vars.append(day_var)  # Add to day shift list
        
        # Night shift checkbox
        night_var = IntVar()
        Checkbutton(scrollable_frame, variable=night_var).grid(row=i, column=3)
        check_vars.append((f"Night {day}", night_var))
        night_shift_vars.append(night_var)  # Add to night shift list
    
    # Pre-select
    existing_selections = selected_prefer_days.get(row_num, [])
    for shift_name, var in check_vars:
        if shift_name in existing_selections:
            var.set(1)

    # Calculate the row number for "Check All" row
    # It should be after all the days, so: len(days_list) + 1
    check_all_row = len(days_list) + 1
    
    # Create the "Check All" label in column 1
    Label(scrollable_frame, text="Check All").grid(row=check_all_row, column=1, sticky="w")
    
    # Create master checkbox variables for day and night
    master_day_var = IntVar()  # Variable for "Check All Day Shifts"
    master_night_var = IntVar()  # Variable for "Check All Night Shifts"
    
    # Function that runs when "Check All Day Shifts" is clicked
    def toggle_all_day_shifts():
        # Get the value: 1 if checked, 0 if unchecked
        value = master_day_var.get()
        # Loop through ALL day shift checkboxes and set them to the same value
        for day_var in day_shift_vars:
            day_var.set(value)  # Set each one to 1 (checked) or 0 (unchecked)
    
    # Function that runs when "Check All Night Shifts" is clicked
    def toggle_all_night_shifts():
        # Get the value: 1 if checked, 0 if unchecked
        value = master_night_var.get()
        # Loop through ALL night shift checkboxes and set them to the same value
        for night_var in night_shift_vars:
            night_var.set(value)  # Set each one to 1 (checked) or 0 (unchecked)
    
    # Create the "Check All" checkboxes with the command parameter
    # command= tells the checkbox what function to run when clicked
    Checkbutton(scrollable_frame, variable=master_day_var, 
                command=toggle_all_day_shifts).grid(row=check_all_row, column=2)
    Checkbutton(scrollable_frame, variable=master_night_var, 
                command=toggle_all_night_shifts).grid(row=check_all_row, column=3)

    def save_selection():
        selected_shifts = [shift for shift, var in check_vars if var.get() == 1]
        # Update the dict (mutates the original dict in main.py)
        selected_prefer_days[row_num] = selected_shifts
        # Update button text
        for row_widgets in worker_rows:
            if row_widgets['row_num'] == row_num:
                num = len(selected_shifts)
                row_widgets['prefer_button'].config(
                    text=f"Select ({num})" if num > 0 else "Select"
                )
                break
        # Show message (exactly like the other popups)
        if selected_shifts:
            error_label.config(
                text=f"Worker prefers to work on these shifts: {', '.join(selected_shifts)}"
            )
        else:
            error_label.config(text="")
        popup.destroy()

    Button(scrollable_frame, text="Save Selection", command=save_selection)\
        .grid(row=check_all_row + 1, column=0, columnspan=4)
    
def cannot_count(cannot_popup_inputs):

    root = cannot_popup_inputs["give_root"]
    days_list = cannot_popup_inputs["give_days_list"]
    starting_weekday = cannot_popup_inputs["give_starting_weekday"]
    worker_rows = cannot_popup_inputs["give_worker_rows"]
    row_num = cannot_popup_inputs["give_row_num"]
    selected_cannot_days = cannot_popup_inputs["give_selected_cannot_days"]
    error_label = cannot_popup_inputs["give_error_label"]

    popup = Toplevel(root)
    popup.title("Select Days Cannot Work")

    # Create a canvas and scrollbar for scrolling
    canvas = Canvas(popup)
    scrollbar = Scrollbar(popup, orient="vertical", command=canvas.yview)
    scrollable_frame = Frame(canvas)

    scrollable_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
    )
    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    def on_popup_mousewheel(event):
        canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")
    
    popup.bind("<MouseWheel>", on_popup_mousewheel)
    canvas.bind("<MouseWheel>", on_popup_mousewheel)

    Label(scrollable_frame, text="Day").grid(row=0, column=0)
    Label(scrollable_frame, text="Weekday").grid(row=0, column=1)
    Label(scrollable_frame, text="Day Shift").grid(row=0, column=2)
    Label(scrollable_frame, text="Night Shift").grid(row=0, column=3)
    
    day_names = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
    check_vars = []
    
    # Create separate lists to hold ONLY day shift variables and ONLY night shift variables
    day_shift_vars = []  # Empty basket for day shift checkboxes
    night_shift_vars = []  # Empty basket for night shift checkboxes
    
    for i, day in enumerate(days_list, start=1):
        weekday = (starting_weekday + (day - 1)) % 7
        day_name = day_names[weekday]
        Label(scrollable_frame, text=str(day)).grid(row=i, column=0)
        Label(scrollable_frame, text=day_name).grid(row=i, column=1)
        
        # Day shift checkbox
        day_var = IntVar()
        Checkbutton(scrollable_frame, variable=day_var).grid(row=i, column=2)
        check_vars.append((f"Day {day}", day_var))
        day_shift_vars.append(day_var)  # Add to day shift list
        
        # Night shift checkbox
        night_var = IntVar()
        Checkbutton(scrollable_frame, variable=night_var).grid(row=i, column=3)
        check_vars.append((f"Night {day}", night_var))
        night_shift_vars.append(night_var)  # Add to night shift list

    # Pre-select checkboxes based on existing selections for this row_num
    existing_selections = selected_cannot_days.get(row_num, [])
    for shift_name, var in check_vars:
        if shift_name in existing_selections:
            var.set(1)

    # Calculate the row number for "Check All" row
    # It should be after all the days, so: len(days_list) + 1
    check_all_row = len(days_list) + 1
    
    # Create the "Check All" label in column 1
    Label(scrollable_frame, text="Check All").grid(row=check_all_row, column=1, sticky="w")
    
    # Create master checkbox variables for day and night
    master_day_var = IntVar()  # Variable for "Check All Day Shifts"
    master_night_var = IntVar()  # Variable for "Check All Night Shifts"
    
    # Function that runs when "Check All Day Shifts" is clicked
    def toggle_all_day_shifts():
        # Get the value: 1 if checked, 0 if unchecked
        value = master_day_var.get()
        # Loop through ALL day shift checkboxes and set them to the same value
        for day_var in day_shift_vars:
            day_var.set(value)  # Set each one to 1 (checked) or 0 (unchecked)
    
    # Function that runs when "Check All Night Shifts" is clicked
    def toggle_all_night_shifts():
        # Get the value: 1 if checked, 0 if unchecked
        value = master_night_var.get()
        # Loop through ALL night shift checkboxes and set them to the same value
        for night_var in night_shift_vars:
            night_var.set(value)  # Set each one to 1 (checked) or 0 (unchecked)
    
    # Create the "Check All" checkboxes with the command parameter
    # command= tells the checkbox what function to run when clicked
    Checkbutton(scrollable_frame, variable=master_day_var, 
                command=toggle_all_day_shifts).grid(row=check_all_row, column=2)
    Checkbutton(scrollable_frame, variable=master_night_var, 
                command=toggle_all_night_shifts).grid(row=check_all_row, column=3)

    def save_selection():
        selected_shifts = [shift for shift, var in check_vars if var.get() == 1] # Creates a list: [Night 5, Day 7] 
        selected_cannot_days[row_num] = selected_shifts
        
        # Update button text
        for row_widgets in worker_rows:
            if row_widgets['row_num'] == row_num:
                num = len(selected_shifts)
                row_widgets['cannot_button'].config(text=f"Select ({num})" if num > 0 else "Select")
                break
        if selected_shifts:
            error_label.config(text=f"Worker cannot work on these shifts: {', '.join(selected_shifts)}")
        else:
            error_label.config(text="")
        popup.destroy()

    # Update the "Save Selection" button row to be AFTER the "Check All" row
    Button(scrollable_frame, text="Save Selection", 
           command=save_selection).grid(row=check_all_row + 1, column=0, columnspan=4)
    
def prefer_unit_count(prefer_unit_popup_inputs):

    root = prefer_unit_popup_inputs["give_root"]
    days_list = prefer_unit_popup_inputs["give_days_list"]
    starting_weekday = prefer_unit_popup_inputs["give_starting_weekday"]
    worker_rows = prefer_unit_popup_inputs["give_worker_rows"]
    row_num = prefer_unit_popup_inputs["give_row_num"]
    selected_manual_days = prefer_unit_popup_inputs["give_selected_manual_days"]
    error_label = prefer_unit_popup_inputs["give_error_label"]
    units_list = prefer_unit_popup_inputs["give_units_list"]
    selected_units = prefer_unit_popup_inputs["give_selected_units"]

    popup = Toplevel(root)
    popup.title("Select Units")

    # Create a canvas and scrollbar for scrolling
    canvas = Canvas(popup)
    scrollbar = Scrollbar(popup, orient="vertical", command=canvas.yview)
    scrollable_frame = Frame(canvas)

    scrollable_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
    )
    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    def on_popup_mousewheel(event):
        canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    popup.bind("<MouseWheel>", on_popup_mousewheel)
    canvas.bind("<MouseWheel>", on_popup_mousewheel)

    Label(scrollable_frame, text=str("List of units")).grid(row=0, column=0)

    unit_vars = [] # Will be used for pre-selecting and saving selection

    for i, unit in enumerate(units_list, start=1):
        Label(scrollable_frame, text=str(unit)).grid(row=i, column=0)
        # Unit checkbox
        unit_var = IntVar()
        Checkbutton(scrollable_frame, variable=unit_var).grid(row=i, column=2)
        unit_vars.append((f"{unit}", unit_var))

    # Pre-select checkboxes based on existing selections for this row_num
    existing_selections = selected_units.get(row_num, [])
    for unit, var in unit_vars: # Go through all units vars
        if unit in existing_selections: # If "Internal Medicine" is in existing_selections, var.set(1) for "Internal Medicine"
            var.set(1)

    check_all_row = len(units_list)

    def save_selection():
        selected_units_list = [unit for unit, var in unit_vars if var.get() == 1] # Gives ["Cardiology", "Oncology"]
        selected_units[row_num] = selected_units_list
        # Update button text
        for row_widgets in worker_rows:
            if row_widgets['row_num'] == row_num:
                num = len(selected_units_list)
                row_widgets['prefer_unit_button'].config(text=f"Select ({num})" if num > 0 else "Select")
                break
        if selected_units_list:
            error_label.config(text=f"Worker prefers these units: {', '.join(selected_units_list)}")
        else:
            error_label.config(text="")
        popup.destroy()

    Button(scrollable_frame, text="Save Selection", command=save_selection).grid(row=check_all_row + 1, column=0)

def manual_count(manual_popup_inputs):

    root = manual_popup_inputs["give_root"]
    days_list = manual_popup_inputs["give_days_list"]
    starting_weekday = manual_popup_inputs["give_starting_weekday"]
    worker_rows = manual_popup_inputs["give_worker_rows"]
    row_num = manual_popup_inputs["give_row_num"]
    selected_manual_days = manual_popup_inputs["give_selected_manual_days"]
    error_label = manual_popup_inputs["give_error_label"]
    units_list = manual_popup_inputs["give_units_list"]

    popup = Toplevel(root)
    popup.title("Select Manual Shifts")

    # Create a canvas and scrollbar for scrolling
    canvas = Canvas(popup)
    scrollbar = Scrollbar(popup, orient="vertical", command=canvas.yview)
    scrollable_frame = Frame(canvas)

    scrollable_frame.bind(
        "<Configure>",
        lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
    )
    canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
    canvas.configure(yscrollcommand=scrollbar.set)

    canvas.pack(side="left", fill="both", expand=True)
    scrollbar.pack(side="right", fill="y")

    def on_popup_mousewheel(event):
        canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")

    popup.bind("<MouseWheel>", on_popup_mousewheel)
    canvas.bind("<MouseWheel>", on_popup_mousewheel)

    Label(scrollable_frame, text="Day").grid(row=0, column=0)
    Label(scrollable_frame, text="Weekday").grid(row=0, column=1)
    Label(scrollable_frame, text="Day Shift").grid(row=0, column=2)
    Label(scrollable_frame, text="Night Shift").grid(row=0, column=3)

    day_names = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]
    check_vars = []

    # Create separate lists to hold ONLY day shift variables and ONLY night shift variables
    day_shift_vars = []  # Empty basket for day shift checkboxes
    night_shift_vars = []  # Empty basket for night shift checkboxes

    for i, day in enumerate(days_list, start=1):
        weekday = (starting_weekday + (day - 1)) % 7
        day_name = day_names[weekday]
        Label(scrollable_frame, text=str(day)).grid(row=i, column=0)
        Label(scrollable_frame, text=day_name).grid(row=i, column=1)
        
        # Day shift checkbox
        day_var = IntVar()
        Checkbutton(scrollable_frame, variable=day_var).grid(row=i, column=2)
        check_vars.append((f"Day {day}", day_var))
        day_shift_vars.append(day_var)  # Add to day shift list
        
        # Night shift checkbox
        night_var = IntVar()
        Checkbutton(scrollable_frame, variable=night_var).grid(row=i, column=3)
        check_vars.append((f"Night {day}", night_var))
        night_shift_vars.append(night_var)  # Add to night shift list

    # Pre-select checkboxes based on existing selections for this row_num
    existing_selections = selected_manual_days.get(row_num, [])
    unit_var = StringVar(value=units_list[0] if units_list else "")  # Default to first unit or empty
    for unit in units_list: # Go through all units to find the shifts like "Day 5 Cardiology"
        for shift_name, var in check_vars: # Go through shift_name(s) in check_vars like "Day 5"
            shift_full_name = f"{shift_name} {unit}" # Set shift_name with unit "Day 5 Cardiology"
            if shift_full_name in existing_selections: # If "Day 5 Cardiology" is in existing_selections, var.set(1) for "Day 5"
                var.set(1)
                # Set the OptionMenu to show the unit
                unit_var.set(unit)

    # Calculate the row number for "Check All" row
    # It should be after all the days, so: len(days_list) + 1
    check_all_row = len(days_list) + 1

    # Create the "Check All" label in column 1
    Label(scrollable_frame, text="Check All").grid(row=check_all_row, column=1, sticky="w")

    # Create master checkbox variables for day and night
    master_day_var = IntVar()  # Variable for "Check All Day Shifts"
    master_night_var = IntVar()  # Variable for "Check All Night Shifts"

    # Function that runs when "Check All Day Shifts" is clicked
    def toggle_all_day_shifts():
        # Get the value: 1 if checked, 0 if unchecked
        value = master_day_var.get()
        # Loop through ALL day shift checkboxes and set them to the same value
        for day_var in day_shift_vars:
            day_var.set(value)  # Set each one to 1 (checked) or 0 (unchecked)

    # Function that runs when "Check All Night Shifts" is clicked
    def toggle_all_night_shifts():
        # Get the value: 1 if checked, 0 if unchecked
        value = master_night_var.get()
        # Loop through ALL night shift checkboxes and set them to the same value
        for night_var in night_shift_vars:
            night_var.set(value)  # Set each one to 1 (checked) or 0 (unchecked)

    # Create the "Check All" checkboxes with the command parameter
    # command= tells the checkbox what function to run when clicked
    Checkbutton(scrollable_frame, variable=master_day_var,
                command=toggle_all_day_shifts).grid(row=check_all_row, column=2)
    Checkbutton(scrollable_frame, variable=master_night_var,
                command=toggle_all_night_shifts).grid(row=check_all_row, column=3)

    # Hospital Unit Checks

    Label(scrollable_frame, text="Hospital Unit").grid(row=check_all_row+1, column=1, sticky="w")
    OptionMenu(scrollable_frame, unit_var, *units_list).grid(row=check_all_row+1, column=2, columnspan=2) # sticky="w"
    
    def save_selection():
        selected_unit = unit_var.get().strip()
        selected_shifts = [shift for shift, var in check_vars if var.get() == 1] # Gives ["Day 5", "Night 7"]
        selected_shifts = [f"{shift} {selected_unit}" for shift in selected_shifts] # Gives ["Day 5 Cardiology", ...]

        selected_manual_days[row_num] = selected_shifts
        # Update button text
        for row_widgets in worker_rows:
            if row_widgets['row_num'] == row_num:
                num = len(selected_shifts)
                row_widgets['manual_button'].config(text=f"Select ({num})" if num > 0 else "Select")
                break
        if selected_shifts:
            error_label.config(text=f"Worker manually assigned to these shifts: {', '.join(selected_shifts)}")
        else:
            error_label.config(text="")
        popup.destroy()

    Button(scrollable_frame, text="Save Selection", command=save_selection).grid(row=check_all_row + 2, column=0, columnspan=4)
