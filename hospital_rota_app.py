import sys # Easy way for Python to restart
import pulp  # Bring in the PuLP toolbox.
import time # Used to show how much time it takes to solve the rota
import calendar  # Bring in the calendar toolbox for month days.
import os # Bring in interaction with Windows/Apple/Linux
import json # Bring in functionality of saving/loading javascript object notation - data-interchange format
from tkinter import *  # Bring in the Tkinter toolbox for the window (GUI).
from tkinter import filedialog

# App split into Part I: Tkinter GUI part and Part II: Python app

# Global variables: IMPORTANT!!!
year = None
month = None
holiday_days = [] # Empty variable, so "Make shifts" works even if Holidays saved nothing
shifts_list = []
workers_list = []
selected_cannot_days = {}  # Changed to dict to store per row_num
selected_prefer_days = {}  # Dict to store preferred days per row_num
selected_manual_days = {}  # Dict to store manual days per row_num

# Global variables: constant
day_names = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]  # Group of day names.
# PuLP settings points
points_filled = 100
points_preferred = 5
points_spacing = -1
spacing_days_threshold = 5  # How many days apart triggers the penalty
points_24hr = -10
# The main window (like the car's dashboard).

root = Tk()  # Make the window box.
#root.geometry("975x450")  # Set predetermined window size (width x height)
root.title("Hospital Rota App - 'Riaukapp' ")  # Name on top.

# Label inside frame.
Label(root, text="Enter the year (e.g., 2026):").pack()  # Text, side=LEFT for horizontal.

# Frame for year row – like a shelf for horizontal.
year_frame = Frame(root)  # Make small box (Frame) inside root.
year_frame.pack()  # Put the shelf on the window.

Label(root, text="Enter the month (1-12, e.g., 1 for January):").pack()  # Text label.

# Type box inside frame.
year_entry = Entry(year_frame)  # Type box in frame.
year_entry.pack(side=LEFT)  # Next to label.

# Function for save year (like button press).
def save_year():
    global year  # Use the year box outside.
    year_input = year_entry.get().strip()  # Get from type box and strip whitespace.
    try:
        year = int(year_input)
        if year < 1900 or year > 2100:
            error_label.config(text="Bad year – 1900-2100.")  # Show error.
        else:
            current_year_label.config(text="Current Year: " + str(year))  # Update show.
            error_label.config(text="")  # Clear error.
    except ValueError:
        error_label.config(text="Not a number!")

# Button inside frame.
Button(year_frame, text="Save", command=save_year).pack(side=LEFT)  # Button next.

# Show current year.
current_year_label = Label(year_frame, text="Current Year: None")  # Show label.
current_year_label.pack(side=LEFT)

month_frame = Frame(root)
month_frame.pack()

# Label and type box for month.
month_entry = Entry(month_frame)  # Type box.
month_entry.pack(side=LEFT)

def save_month():  # Function for the button.
    global month, num_days, starting_weekday, days_list, shifts_list  # Use these boxes outside – add shifts_list.
    month_input = month_entry.get()  # Get from type box.
    try:
        month = int(month_input)
        if month < 1 or month > 12:
            error_label.config(text="Please enter a month between 1 and 12.")  # Error.
            return  # Stop early if bad.
        # Calculate details.
        month_details = calendar.monthrange(year, month)
        starting_weekday = month_details[0]  # Weekday for Day 1.
        num_days = month_details[1]  # Days in month.
        # Make the days list
        days_list = []  # Empty list to hold days.
        for day in range(1, num_days + 1):  # Loop from 1 to num_days +1 (to include last).
            days_list.append(day)  # Add day to list.
        # Show the details.
        day_names = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]  # Group of day names.
        # Show the days list.
        month_name_list = ["January", "February", "March", "April", "May", "June", "July", "September", "October", "November", "December"]
        month_name = month_name_list[month-1]
        current_month_label.config(text="Current Month: " + str(month_name)) # Update month
        error_label.config(text=f"{month_name} has {len(days_list)} days, start of the weekday: {day_names[starting_weekday]}")  # Update result.
    except ValueError:
        error_label.config(text="That's not a number! Please try again.")  # Error.

Button(month_frame, text="Save", command=save_month).pack(side=LEFT)  # Button, click runs save_month.

current_month_label = Label(month_frame, text="Current Month: None")
current_month_label.pack(side=LEFT)

def show_cannot_popup(row_num):
    global selected_cannot_days
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
    for i, day in enumerate(days_list, start=1):
        weekday = (starting_weekday + (day - 1)) % 7
        day_name = day_names[weekday]
        Label(scrollable_frame, text=str(day)).grid(row=i, column=0)
        Label(scrollable_frame, text=day_name).grid(row=i, column=1)
        # Day shift checkbox
        day_var = IntVar()
        Checkbutton(scrollable_frame, variable=day_var).grid(row=i, column=2)
        check_vars.append((f"Day {day}", day_var))
        # Night shift checkbox
        night_var = IntVar()
        Checkbutton(scrollable_frame, variable=night_var).grid(row=i, column=3)
        check_vars.append((f"Night {day}", night_var))

    # Pre-select checkboxes based on existing selections for this row_num
    existing_selections = selected_cannot_days.get(row_num, [])
    for shift_name, var in check_vars:
        if shift_name in existing_selections:
            var.set(1)

    def save_selection():
        global selected_cannot_days
        selected_shifts = [shift for shift, var in check_vars if var.get() == 1]
        selected_cannot_days[row_num] = selected_shifts
        # Update button text
        for row_widgets in worker_rows:
            if row_widgets['row_num'] == row_num:
                num = len(selected_shifts)
                row_widgets['cannot_button'].config(text=f"Select Days ({num})" if num > 0 else "Select Days")
                break
        if selected_shifts:
            error_label.config(text=f"Worker cannot work on these shifts: {', '.join(selected_shifts)}")
        else:
            error_label.config(text="")
        popup.destroy()

    Button(scrollable_frame, text="Save Selection", command=save_selection).grid(row=len(days_list)+1, column=0, columnspan=4)

def show_prefer_popup(row_num):
    global selected_prefer_days
    popup = Toplevel(root)
    popup.title("Select Days Prefer Work")

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
    for i, day in enumerate(days_list, start=1):
        weekday = (starting_weekday + (day - 1)) % 7
        day_name = day_names[weekday]
        Label(scrollable_frame, text=str(day)).grid(row=i, column=0)
        Label(scrollable_frame, text=day_name).grid(row=i, column=1)
        # Day shift checkbox
        day_var = IntVar()
        Checkbutton(scrollable_frame, variable=day_var).grid(row=i, column=2)
        check_vars.append((f"Day {day}", day_var))
        # Night shift checkbox
        night_var = IntVar()
        Checkbutton(scrollable_frame, variable=night_var).grid(row=i, column=3)
        check_vars.append((f"Night {day}", night_var))

    # Pre-select checkboxes based on existing selections for this row_num
    existing_selections = selected_prefer_days.get(row_num, [])
    for shift_name, var in check_vars:
        if shift_name in existing_selections:
            var.set(1)

    def save_selection():
        global selected_prefer_days
        selected_shifts = [shift for shift, var in check_vars if var.get() == 1]
        selected_prefer_days[row_num] = selected_shifts
        # Update button text
        for row_widgets in worker_rows:
            if row_widgets['row_num'] == row_num:
                num = len(selected_shifts)
                row_widgets['prefer_button'].config(text=f"Select Days ({num})" if num > 0 else "Select Days")
                break
        if selected_shifts:
            error_label.config(text=f"Worker prefers to work on these shifts: {', '.join(selected_shifts)}")
        else:
            error_label.config(text="")
        popup.destroy()

    Button(scrollable_frame, text="Save Selection", command=save_selection).grid(row=len(days_list)+1, column=0, columnspan=4)

def show_manual_popup(row_num):
    global selected_manual_days
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
    for i, day in enumerate(days_list, start=1):
        weekday = (starting_weekday + (day - 1)) % 7
        day_name = day_names[weekday]
        Label(scrollable_frame, text=str(day)).grid(row=i, column=0)
        Label(scrollable_frame, text=day_name).grid(row=i, column=1)
        # Day shift checkbox
        day_var = IntVar()
        Checkbutton(scrollable_frame, variable=day_var).grid(row=i, column=2)
        check_vars.append((f"Day {day}", day_var))
        # Night shift checkbox
        night_var = IntVar()
        Checkbutton(scrollable_frame, variable=night_var).grid(row=i, column=3)
        check_vars.append((f"Night {day}", night_var))

    # Pre-select checkboxes based on existing selections for this row_num
    existing_selections = selected_manual_days.get(row_num, [])
    for shift_name, var in check_vars:
        if shift_name in existing_selections:
            var.set(1)

    def save_selection():
        global selected_manual_days
        selected_shifts = [shift for shift, var in check_vars if var.get() == 1]
        selected_manual_days[row_num] = selected_shifts
        # Update button text
        for row_widgets in worker_rows:
            if row_widgets['row_num'] == row_num:
                num = len(selected_shifts)
                row_widgets['manual_button'].config(text=f"Select Manual ({num})" if num > 0 else "Select Manual")
                break
        if selected_shifts:
            error_label.config(text=f"Worker manually assigned to these shifts: {', '.join(selected_shifts)}")
        else:
            error_label.config(text="")
        popup.destroy()

    Button(scrollable_frame, text="Save Selection", command=save_selection).grid(row=len(days_list)+1, column=0, columnspan=4)

Label(root, text="Enter public holiday days, separated by comma (e.g., 24,25,26) or leave blank:").pack()  # Text label.

holiday_frame = Frame(root)
holiday_frame.pack()

# Label and type box for holidays.
holiday_entry = Entry(holiday_frame)  # Type box.
holiday_entry.pack(side=LEFT)

def save_holidays():  # Function for the button.
    global holiday_days  # Use the holiday_days box outside.
    holiday_input = holiday_entry.get()  # Get from type box.
    holiday_days = []  # Start empty.
    if holiday_input.strip() == "":  # If blank.
        holiday_days = []
        holidays_label.config(text="Holiday List: None")
        error_label.config(text="")  # Clear error.
        return  # Done.
    parts = holiday_input.split(",")  # Cut at commas.
    try:
        holiday_days = [int(part.strip()) for part in parts]  # Turn to numbers.
        invalid_days = []
        for h_day in holiday_days[:]:  # Copy to avoid remove issues.
            if h_day < 1 or h_day > num_days:  # Out of range.
                invalid_days.append(h_day)
                holiday_days.remove(h_day)
        if invalid_days:
            error_label.config(text=f"Error: These days are invalid (must be 1-{num_days}): {invalid_days}")  # Show error.
            return  # Stop early.
        holidays_label.config(text="Holiday List: " + str(holiday_days))
        error_label.config(text="")  # Clear if good.
    except ValueError:
        error_label.config(text="Error: Some entries weren't numbers. Please try again.")  # Error.

Button(holiday_frame, text="Save", command=save_holidays).pack(side=LEFT)  # Button, click runs save_holidays.

holidays_label = Label(holiday_frame, text="Holiday List: None")
holidays_label.pack(side=RIGHT)

make_shifts_button_frame = Frame(root)
make_shifts_button_frame.pack()

add_worker_button_frame = Frame(root)
add_worker_button_frame.pack()

# Button to add row.

# Frame to hold the scrollable worker area (like a shelf for the canvas)
worker_container = Frame(root)
#worker_container.pack(fill="both", expand=True)  # Fill the space, can grow
worker_container.pack()

# Create the canvas (drawing board)
global worker_canvas, worker_scrollbar, worker_inner_frame
worker_canvas = Canvas(worker_container, width=830)  # Height=200 pixels – change if you want taller/shorter
worker_canvas.pack(side=LEFT, fill="both")  # Put on left, fill space

# Create the scrollbar and link it to the canvas
worker_scrollbar = Scrollbar(worker_container, orient="vertical", command=worker_canvas.yview)
worker_scrollbar.pack(side=RIGHT, fill="y")  # Put on right, vertical

# Link canvas back to scrollbar (so it knows when to show the slider)
worker_canvas.configure(yscrollcommand=worker_scrollbar.set)

# Create the inner frame (where worker rows will go)
worker_inner_frame = Frame(worker_canvas)
worker_canvas.create_window((0, 0), window=worker_inner_frame, anchor="nw")  # Put the frame at top-left of canvas

# Mouse wheel
def on_mouse_wheel(event):
    worker_canvas.yview_scroll(int(-1 * (event.delta / 120)), "units")  # Scroll up/down with wheel

root.bind("<MouseWheel>", on_mouse_wheel)  # Link wheel to function

# Make the canvas scroll when the inner frame grows
def update_scroll_region(event):
    worker_canvas.configure(scrollregion=worker_canvas.bbox("all"))  # Keep this – updates scroll area

    # Ask inner_frame how tall it needs to be (reqheight = "required height")
    required_height = worker_inner_frame.winfo_reqheight()  # This is like measuring the stack of rows

    # Decide the height: Min of required or 200 (the max)
    max_height = 300  # Your max – change if you want
    new_height = min(required_height, max_height)  # min() picks the smaller one

    # Set the canvas height to that
    worker_canvas.config(height=new_height)  # Update it!

    # If required > max, scrollbar will show automatically – no extra code needed

def update_inner_width(event):
    worker_inner_frame.config(width=event.width)  # Set width to match canvas

worker_canvas.bind("<Configure>", update_inner_width)  # Run this when canvas size changes

worker_inner_frame.bind("<Configure>", update_scroll_region)  # Run this function when inner frame changes size

def make_shifts():  # New function for making shifts list.
    global shifts_list, holiday_days  # Use the shifts_list box outside.
    shifts_list = []  # Empty list for shifts.
    shift_types = ["Day", "Night"]  # Group of types.
    day_names = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]  # Group of day names.
    for day in days_list:  # Loop each day.
        weekday = (starting_weekday + (day - 1)) % 7  # Calculate day name ( %7 like clock wrap).
        tags = [day_names[weekday]]  # Add day name tag.
        if weekday in [5, 6]:  # Sat/Sun – weekend.
            tags.append("Weekend")  # Add tag.
        if day in holiday_days:  # If holiday.
            tags.append("Public holiday")  # Add tag.
        for shift_type in shift_types:  # Loop Day then Night.
            # Check if this is a Monday-Friday Day shift (and not a holiday) – if both true, exclude (skip)
            if weekday in [0, 1, 2, 3, 4] and shift_type == "Day" and day not in holiday_days:
                continue  # Skip adding this shift – it's excluded.
            shift_name = f"{shift_type} {day}"  # Make name.
            shift_dict = {  # Make the shift box.
                "name": shift_name,
                "type": shift_type,
                "tags": tags,
                "assigned_worker": None
            }
            shifts_list.append(shift_dict)  # Add to list.
    error_label.config(text="Shifts made: " + str(len(shifts_list)))  # Update the label with count.
    # Show the shifts list nicely in terminal, for debugging
    print("Your shifts list (with tags, types):")
    for shift in shifts_list:  # Loop to print each one.
        print(f"Shift: {shift['name']}, Tags: {shift['tags']}, Type: {shift['type']}, Assigned: {shift['assigned_worker']}")

Button(make_shifts_button_frame, text="Make Shifts", command=make_shifts, width=10).pack()  # Button, click runs make_shifts.

# Box to show the shifts list.
#shifts_label = Label(root, text="Shifts made: None")  # Start None.
#shifts_label.pack()  # Put on window.

# Frame for workers – like a big shelf for the table.
workers_frame = Frame(worker_inner_frame)  # NEW: Pack to inner_frame
workers_frame.pack(fill="x")  # NEW: Fill horizontal, stack vertical

# Configure columns to have fixed width of 50
for col in range(10):
    workers_frame.columnconfigure(col, minsize=10, weight=0)

# Global headers for columns.
Label(workers_frame, text="Name", width=10, padx=2, pady=2).grid(row=0, column=0, sticky="ew", padx=2, pady=2)
Label(workers_frame, text="Shift Range", width=10, padx=2, pady=2).grid(row=0, column=1, sticky="ew", padx=2, pady=2)
Label(workers_frame, text="Cannot Days", width=10, padx=2, pady=2).grid(row=0, column=2, sticky="ew", padx=2, pady=2)
Label(workers_frame, text="Prefer Days", width=10, padx=2, pady=2).grid(row=0, column=3, sticky="ew", padx=2, pady=2)
Label(workers_frame, text="Max Wknds", width=10, padx=2, pady=2).grid(row=0, column=4, sticky="ew", padx=2, pady=2)
Label(workers_frame, text="Max 24hr", width=10, padx=2, pady=2).grid(row=0, column=5, sticky="ew", padx=2, pady=2)
Label(workers_frame, text="Manual Shifts", width=10, padx=2, pady=2).grid(row=0, column=6, sticky="ew", padx=2, pady=2)
Label(workers_frame, text="Save", width=10, padx=2, pady=2).grid(row=0, column=7, sticky="ew", padx=2, pady=2)
Label(workers_frame, text="Assign", width=10, padx=2, pady=2).grid(row=0, column=8, sticky="ew", padx=2, pady=2)
Label(workers_frame, text="Delete", width=10, padx=2, pady=2).grid(row=0, column=9, sticky="ew", padx=2, pady=2)

# List to hold worker rows (for delete).
worker_rows = []  # Empty list to store widget references for each row.
worker_row_number = 1  # Start row 1 (after 0 headers)

def add_worker_row():  # Function for "Add Worker" button.
    global worker_row_number, workers_list
    row_num = worker_row_number  # Capture the current row number for this row

    # Column 0: Name box - placed directly in workers_frame.
    name_entry = Entry(workers_frame, width=10)  # Type box.
    name_entry.grid(row=row_num, column=0, sticky="ew", padx=2, pady=2)

    # Column 1: Shift range box.
    range_entry = Entry(workers_frame, width=10)
    range_entry.grid(row=row_num, column=1, sticky="ew", padx=2, pady=2)

    # Column 2: Cannot work button.
    cannot_button = Button(workers_frame, text="Select Days", width=10, command=lambda: show_cannot_popup(row_num))
    cannot_button.grid(row=row_num, column=2, sticky="ew", padx=2, pady=2)

    # Column 3: Prefer button.
    prefer_button = Button(workers_frame, text="Select Days", width=10, command=lambda: show_prefer_popup(row_num))
    prefer_button.grid(row=row_num, column=3, sticky="ew", padx=2, pady=2)

    # Column 4: Max weekends box.
    max_weekends_entry = Entry(workers_frame, width=10)
    max_weekends_entry.grid(row=row_num, column=4, sticky="ew", padx=2, pady=2)

    # Column 5: Max 24-hour box.
    max_24hr_entry = Entry(workers_frame, width=10)
    max_24hr_entry.grid(row=row_num, column=5, sticky="ew", padx=2, pady=2)

    # Column 6: Manual shifts button.
    manual_button = Button(workers_frame, text="Select Manual", width=10, command=lambda: show_manual_popup(row_num))
    manual_button.grid(row=row_num, column=6, sticky="ew", padx=2, pady=2)

    # Column 7: Save Worker button.
    save_button = Button(workers_frame, text="Save", width=10, command=lambda: save_worker(row_num))
    save_button.grid(row=row_num, column=7, sticky="ew", padx=2, pady=2)

    # Column 8: Manual save button.
    manual_save_button = Button(workers_frame, text="Assign Manual", width=10, command=lambda: save_manual(row_num))
    manual_save_button.grid(row=row_num, column=8, sticky="ew", padx=2, pady=2)

    # Column 9: Delete button.
    delete_button = Button(workers_frame, text="Delete", width=10, command=lambda: delete_row(row_num))
    delete_button.grid(row=row_num, column=9, sticky="ew", padx=2, pady=2)

    # Store all widgets for this row for deletion purposes
    row_widgets = {
        'name_entry': name_entry,
        'range_entry': range_entry,
        'cannot_button': cannot_button,
        'prefer_button': prefer_button,
        'max_weekends_entry': max_weekends_entry,
        'max_24hr_entry': max_24hr_entry,
        'manual_button': manual_button,
        'save_button': save_button,
        'manual_save_button': manual_save_button,
        'delete_button': delete_button,
        'row_num': row_num
    }
    worker_rows.append(row_widgets)
    print(f"Current row: {worker_row_number}") #Debugging
    worker_row_number += 1  # Add 1.

    def save_worker(row_num):  # The function – the "recipe" for saving a worker.
        # This runs when you click "Save Worker".
        # Get what you typed from the boxes.
        name = name_entry.get().strip() # Get the name
        range_input = range_entry.get().strip()  # Get shift range.
        max_weekends_input = max_weekends_entry.get().strip()  # Get max weekends.
        max_24hr_input = max_24hr_entry.get().strip()  # Get max 24-hour.

        # Check if shifts_list is created
        global shifts_list
        try:
            if not shifts_list:
                error_label.config(text="Shifts list is not created. Cannot save worker.")
                return
        except NameError:
            error_label.config(text="Shifts list is not created. Cannot save worker.")
            return

        # Unassign all shifts assigned to this worker's name before saving
        for shift in shifts_list:
            if shift["assigned_worker"] == name:
                shift["assigned_worker"] = None

        # Check if name is not blank.
        if name == "":  # If empty.
            error_label.config(text="Error: Name can't be blank.")  # Show error.
            return  # Stop early.

        # Check shift range (like "1-4").
        if range_input == "":  # If blank.
            error_label.config(text="Error: Shift range can't be blank.")  # Error.
            return
        range_parts = range_input.split("-")  # Cut at "-".
        if len(range_parts) != 2:  # Must be 2.
            error_label.config(text="Error: Range like 1-4.")  # Error.
            return
        try:
            min_shifts = int(range_parts[0])
            max_shifts = int(range_parts[1])
            if min_shifts > max_shifts or min_shifts < 0:
                error_label.config(text="Error: Min <= Max, positive.")  # Error.
                return
        except ValueError:
            error_label.config(text="Error: Range not numbers.")  # Error.
            return

        # Cannot work from selected days.
        global selected_cannot_days
        cannot_work = selected_cannot_days.get(row_num, [])  # Get the list for this row_num.

        prefer = selected_prefer_days.get(row_num, [])  # Get the list for this row_num.

        # Max weekends and 24hr – numbers.
        max_weekends = 100  # Default 100.
        if max_weekends_input != "":
            try:
                max_weekends = int(max_weekends_input)
                if max_weekends < 0:
                    error_label.config(text="Error: Max weekends positive or 0.")
                    return
            except ValueError:
                error_label.config(text="Error: Max weekends not number.")
                return

        max_24hr = 100  # Default 100.
        if max_24hr_input != "":
            try:
                max_24hr = int(max_24hr_input)
                if max_24hr < 0:
                    error_label.config(text="Error: Max 24hr positive or 0.")
                    return
            except ValueError:
                error_label.config(text="Error: Max 24hr not number.")
                return

        # ===== Check if worker already exists =====
        # Look through workers_list to find this worker by row number
        existing_worker = None  # Start with None (not found).
        for worker in workers_list:
            if worker["worker_row_number"] == row_num:
                existing_worker = worker  # Found them!
                break  # Stop looking.

        if existing_worker:  # If we found them (updating)
            # Update the existing worker's data
            existing_worker["name"] = name
            existing_worker["shifts_to_fill"] = [min_shifts, max_shifts]
            existing_worker["cannot_work"] = cannot_work
            existing_worker["prefers"] = prefer
            existing_worker["max_weekends"] = max_weekends
            existing_worker["max_24hr"] = max_24hr
            # Check if there were manually selected shifts
            had_manual_shifts = bool(selected_manual_days.get(row_num, []))
            message = f"Worker '{name}' updated!"
            if had_manual_shifts:
                message += " Manually assigned shifts cleared."
            error_label.config(text=message)  # Show success.
        else:  # If not found (new worker)
            # Create new worker dictionary
            worker_dict = {
                "name": name,
                "shifts_to_fill": [min_shifts, max_shifts],
                "cannot_work": cannot_work,
                "prefers": prefer,
                "max_weekends": max_weekends,
                "max_24hr": max_24hr,
                "worker_row_number": row_num
            }
            workers_list.append(worker_dict)  # Add new worker.
            error_label.config(text=f"Worker '{name}' saved!")  # Show success.

        # Print for debugging
        print("Your full worker list:")
        for worker in workers_list:
            print(f"Name: {worker['name']}, shifts: {worker['shifts_to_fill']}, Cannot: {worker['cannot_work']}, Prefers: {worker['prefers']}, Max weekends: {worker['max_weekends']}, Max 24hr: {worker['max_24hr']}, Row number: {worker['worker_row_number']}")

    def save_manual(row_num):
        # Get worker name
        name = name_entry.get().strip()
        if not name:
            error_label.config(text="Error: Worker name required for manual save.")
            return

        # Find worker
        worker = None
        for w in workers_list:
            if w.get("worker_row_number") == row_num:
                worker = w
                break
        if not worker:
            error_label.config(text="Error: Save worker first before manual assignment.")
            return

        # Get selected shifts
        manual_shifts = selected_manual_days.get(row_num, [])
        if not manual_shifts:
            error_label.config(text="No manual shifts selected.")
            return

        successful_assigns = 0
        for shift_name in manual_shifts:
            # Find shift in shifts_list
            shift = None
            for s in shifts_list:
                if s["name"] == shift_name:
                    shift = s
                    break
            if not shift:
                error_label.config(text=f"Error: Shift {shift_name} not found.")
                print("Error: Shift {shift_name} not found.")
                continue
            if shift["assigned_worker"] is not None:
                error_label.config(text=f"Error: Shift {shift_name} already assigned.")
                print("Error: Shift {shift_name} already assigned.")
                continue
            # Assign
            shift["assigned_worker"] = name
            successful_assigns += 1

        if successful_assigns > 0:
            # Update worker's range
            original_min = worker["shifts_to_fill"][0]
            original_max = worker["shifts_to_fill"][1]
            new_min = max(0, original_min - successful_assigns)
            new_max = max(0, original_max - successful_assigns)
            worker["shifts_to_fill"] = [new_min, new_max]
            range_entry.delete(0, END)
            range_entry.insert(0, f"{new_min}-{new_max}")
            error_label.config(text=f"Assigned {successful_assigns} manual shifts to {name}. Updated range to {new_min}-{new_max}.")
            for shift in shifts_list:
                print(shift)
            for worker in workers_list:
                print(worker)
        else:
            error_label.config(text="No shifts assigned.")


def delete_row(row_num):
    # Find the worker's name before deleting
    worker_name = None
    global workers_list, shifts_list, worker_rows
    for worker in workers_list:
        if worker.get("worker_row_number") == row_num:
            worker_name = worker["name"]
            break

    # Unassign all shifts assigned to this worker
    if worker_name:
        for shift in shifts_list:
            if shift["assigned_worker"] == worker_name:
                shift["assigned_worker"] = None

    # Find and destroy all widgets in this row
    for row_widgets in worker_rows:
        if row_widgets['row_num'] == row_num:
            # Destroy each widget
            for key, widget in row_widgets.items():
                if key != 'row_num' and widget.winfo_exists():
                    widget.destroy()
            # Remove from worker_rows list
            worker_rows.remove(row_widgets)
            break

    # Remove the worker from workers_list if it exists
    workers_list = [worker for worker in workers_list if worker.get("worker_row_number") != row_num]
    error_label.config(text=f"Worker '{worker_name}' deleted and all their shifts unassigned!")

def pulp_settings():
    popup = Toplevel(root)
    popup.title("PuLP Settings")

    Label(popup, text="Points for shifts filled:").grid(row=0, column=0, sticky="w")
    filled_entry = Entry(popup)
    filled_entry.insert(0, str(points_filled))
    filled_entry.grid(row=0, column=1)

    Label(popup, text="Points for preferred shifts:").grid(row=1, column=0, sticky="w")
    preferred_entry = Entry(popup)
    preferred_entry.insert(0, str(points_preferred))
    preferred_entry.grid(row=1, column=1)

    Label(popup, text="Points for bad spacing pairs:").grid(row=2, column=0, sticky="w")
    spacing_entry = Entry(popup)
    spacing_entry.insert(0, str(points_spacing))
    spacing_entry.grid(row=2, column=1)

    Label(popup, text="Days apart for spacing penalty:").grid(row=3, column=0, sticky="w")
    spacing_days_entry = Entry(popup)
    spacing_days_entry.insert(0, str(spacing_days_threshold))
    spacing_days_entry.grid(row=3, column=1)

    Label(popup, text="Points for 24-hour shifts:").grid(row=4, column=0, sticky="w")
    hr24_entry = Entry(popup)
    hr24_entry.insert(0, str(points_24hr))
    hr24_entry.grid(row=4, column=1)

    def save_settings():
        global points_filled, points_preferred, points_spacing, spacing_days_threshold, points_24hr
        try:
            points_filled = int(filled_entry.get())
            points_preferred = int(preferred_entry.get())
            points_spacing = int(spacing_entry.get())
            spacing_days_threshold = int(spacing_days_entry.get())
            points_24hr = int(hr24_entry.get())
            error_label.config(text="PuLP settings updated successfully!")
            popup.destroy()
        except ValueError:
            error_label.config(text="Error: All values must be integers.")

    Button(popup, text="Save Settings", command=save_settings).grid(row=5, column=0, columnspan=2)

def save_preferences():
    if year is None:
        error_label.config(text="Year not set.")
        return
    file_path = filedialog.asksaveasfilename(defaultextension=".json", filetypes=[("JSON files", "*.json"), ("All files", "*.*")])
    if file_path:
        data = {"year": year}
        if month is not None:
            data["month"] = month
        data["holiday_days"] = holiday_days
        data["shifts_list"] = shifts_list
        data["workers_list"] = workers_list
        data["selected_cannot_days"] = selected_cannot_days
        data["selected_prefer_days"] = selected_prefer_days
        data["selected_manual_days"] = selected_manual_days
        with open(file_path, 'w') as f:
            json.dump(data, f)
        error_label.config(text="Preferences saved.")

def load_preferences():
    global selected_cannot_days, selected_prefer_days, selected_manual_days
    file_path = filedialog.askopenfilename(filetypes=[("JSON files", "*.json"), ("All files", "*.*")])
    if file_path:
        with open(file_path, 'r') as f:
            data = json.load(f)
        if "year" in data:
            global year
            year = data["year"]
            current_year_label.config(text="Current Year: " + str(year))
            year_entry.delete(0, END)
            year_entry.insert(0, str(year))
        if "month" in data:
            global month
            month = data["month"]
            month_entry.delete(0, END)
            month_entry.insert(0, str(month))
            month_name_list = ["January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"]
            month_name = month_name_list[month-1]
            current_month_label.config(text="Current Month: " + month_name)
            # Since month is loaded, call save_month to set days_list etc.
            save_month()
        if "holiday_days" in data:
            global holiday_days
            holiday_days = data["holiday_days"]
            holiday_entry.delete(0, END)
            if holiday_days:
                holiday_entry.insert(0, ", ".join(map(str, holiday_days)))
            holidays_label.config(text="Holiday List: " + str(holiday_days))
        if "shifts_list" in data:
            global shifts_list
            shifts_list = data["shifts_list"]
        if "workers_list" in data:
            global workers_list
            workers_list = data["workers_list"]
            # Populate selected_cannot_days from workers_list if not in data
            if "selected_cannot_days" not in data:
                selected_cannot_days = {worker["worker_row_number"]: worker.get("cannot_work", []) for worker in workers_list}
            # Clear existing worker rows
            global worker_rows, worker_row_number
            for row_widgets in worker_rows:
                for key, widget in row_widgets.items():
                    if key != 'row_num' and widget.winfo_exists():
                        widget.destroy()
            worker_rows = []
            worker_row_number = 1
            # Now add rows for loaded workers
            for worker in workers_list:
                row_num = worker["worker_row_number"]
                worker_row_number = row_num
                add_worker_row()
            # Now populate the entries
            for worker in workers_list:
                row_num = worker["worker_row_number"]
                for row_widgets in worker_rows:
                    if row_widgets['row_num'] == row_num:
                        row_widgets['name_entry'].delete(0, END)
                        row_widgets['name_entry'].insert(0, worker["name"])
                        row_widgets['range_entry'].delete(0, END)
                        row_widgets['range_entry'].insert(0, f"{worker['shifts_to_fill'][0]}-{worker['shifts_to_fill'][1]}")
                        row_widgets['max_weekends_entry'].delete(0, END)
                        row_widgets['max_weekends_entry'].insert(0, str(worker["max_weekends"]))
                        row_widgets['max_24hr_entry'].delete(0, END)
                        row_widgets['max_24hr_entry'].insert(0, str(worker["max_24hr"]))
                        break
            # Update worker_row_number to max +1
            if workers_list:
                max_row = max(worker["worker_row_number"] for worker in workers_list)
                worker_row_number = max_row + 1
            else:
                worker_row_number = 1
            # Update button texts to show selections
            for row_widgets in worker_rows:
                row_num = row_widgets['row_num']
                num_cannot = len(selected_cannot_days.get(row_num, []))
                row_widgets['cannot_button'].config(text=f"Select Days ({num_cannot})" if num_cannot > 0 else "Select Days")
                num_prefer = len(selected_prefer_days.get(row_num, []))
                row_widgets['prefer_button'].config(text=f"Select Days ({num_prefer})" if num_prefer > 0 else "Select Days")
                num_manual = len(selected_manual_days.get(row_num, []))
                row_widgets['manual_button'].config(text=f"Select Manual ({num_manual})" if num_manual > 0 else "Select Manual")
        if "selected_cannot_days" in data:
            selected_cannot_days = {int(k): v for k, v in data["selected_cannot_days"].items()}
        if "selected_prefer_days" in data:
            selected_prefer_days = {int(k): v for k, v in data["selected_prefer_days"].items()}
        if "selected_manual_days" in data:
            selected_manual_days = {int(k): v for k, v in data["selected_manual_days"].items()}
        # Set worker dicts from loaded selected dicts to ensure consistency
        for worker in workers_list:
            row_num = worker["worker_row_number"]
            worker["cannot_work"] = selected_cannot_days.get(row_num, [])
            worker["prefers"] = selected_prefer_days.get(row_num, [])
            if row_num not in selected_manual_days:
                selected_manual_days[row_num] = []
        error_label.config(text="Preferences loaded.")
    print(f"Debugging. Year: {year}, month {month}, Holiday list: {holiday_days}, Shift list length {len(shifts_list)}")
    for worker in workers_list:
        print(worker)
    print(selected_cannot_days)

# ----------------------------------------
#             PuLP Solve
# ----------------------------------------

def create_rota():
    # Part 6: PuLP to assign workers to empty shifts.

    # First, make assignments dict from shifts_list (shift_name: worker or None).
    assignments = {}
    for shift in shifts_list:
        assignments[shift["name"]] = shift["assigned_worker"]

    # Get empty shifts and empty weekend shifts (None).
    empty_shifts = [name for name in assignments if assignments[name] is None]
    empty_weekend_shifts = []  # Empty list to hold weekend empty names.
    for shift in shifts_list:  # Loop all shifts.
        if shift["name"] in empty_shifts and "Weekend" in shift["tags"]:  # If empty and weekend tag.
            empty_weekend_shifts.append(shift["name"])  # Add name.

    # Get worker names.
    workers = [w["name"] for w in workers_list]

    # Pre-find bad pairs for these rules: bad night to day, bad adjacent nights, bad adjacent days, 24-hour shifts
    bad_night_to_day_pairs = []  # For Night then Day +1.
    bad_adjacent_nights_pairs = []  # For Night then Night +1.
    bad_adjacent_days_pairs = []  # For Day then Day +1.
    twenty_four_hour_shift_pairs = [] # For 24-hour shifts: Day 1 + Night 1

    sorted_shift_names = sorted(empty_shifts, key=lambda s: int(s.split(" ")[1]))  # Sort empty shifts by day.
    for i in range(len(sorted_shift_names)):  # Outer loop.
        for j in range(i+1, len(sorted_shift_names)):  # Inner for pairs after i.
            current = sorted_shift_names[i]
            next_shift = sorted_shift_names[j]
            current_day = int(current.split(" ")[1])
            next_day = int(next_shift.split(" ")[1])
            if next_day == current_day + 1:  # Only if consecutive days – common for all.
                if "Night" in current and "Day" in next_shift:
                    bad_night_to_day_pairs.append((current, next_shift))
                if "Night" in current and "Night" in next_shift:
                    bad_adjacent_nights_pairs.append((current, next_shift))
                if "Day" in current and "Day" in next_shift:
                    bad_adjacent_days_pairs.append((current, next_shift))
            if next_day == current_day:
                twenty_four_hour_shift_pairs.append((current, next_shift))
            
    # Spacing <5 days apart. Check all shift pairs (full rota). If a pair is <5 days apart, add it to the bad list, but only if a shift is empty.
    all_shift_names = sorted([shift["name"] for shift in shifts_list], key=lambda s: int(s.split(" ")[1]))
    bad_spacing_pairs = []  # Empty list.
    for i in range(len(all_shift_names)):  # Outer loop.
        for j in range(i+1, len(all_shift_names)):  # Inner.
            shift1 = all_shift_names[i]  # First.
            shift2 = all_shift_names[j]  # Second.
            day1 = int(shift1.split(" ")[1])
            day2 = int(shift2.split(" ")[1])
            #Add if I want 5 SHIFTS apart, not days: if j - i - 1 < 5:  # Close in SHIFTS (position j - i -1 <5).
            if day2 - day1 < spacing_days_threshold:  # Close - using configurable threshold.
                if shift1 in empty_shifts and shift2 in empty_shifts:  # Only if both empty.
                    bad_spacing_pairs.append((shift1, shift2))  # Add.

    worker_prefers = {w["name"]: w["prefers"] for w in workers_list}
    max_24hr = {w["name"]: w["max_24hr"] for w in workers_list}  # { "Mantas": 2, "Tadas": 3 } – short way to make from list.
    max_weekends = {w["name"]: w["max_weekends"] for w in workers_list}
    worker_cannot = {w["name"]: w["cannot_work"] for w in workers_list}  # Box with worker: their "NO" list.

    # If no empty shifts or no workers, skip.
    if not empty_shifts or not workers:
        print("No empty shifts or no workers – nothing to assign.")
        error_label.config(text="No empty shifts or no workers – nothing to assign.")
    else:
        # Make the PuLP problem.
        prob = pulp.LpProblem("Rota_Assignment", pulp.LpMaximize)  # Maximize points.

        # Variables: Yes/no for worker on shift.
        assign_vars = pulp.LpVariable.dicts("Assign", (workers, empty_shifts), 0, 1, pulp.LpBinary)  # 0 or 1 (no/yes).
        # Define a binary var for if w does 24-hour on that day.
        twenty_four_vars = pulp.LpVariable.dicts("24hr", (workers, twenty_four_hour_shift_pairs), 0, 1, pulp.LpBinary)  # 0 or 1 if does both.
        # Extra boxes for if w assigned to both in bad spacing pair (1 if yes, 0 no).
        spacing_var = pulp.LpVariable.dicts("SpacingBad", (workers, bad_spacing_pairs), 0, 1, pulp.LpBinary)  # Yes/no for each w and pair.

        # Goal: Maximize goals: points for filled, preferred, spacing, 24hr.
        prob += (
        points_filled * pulp.lpSum(assign_vars[w][shift] for w in workers for shift in empty_shifts)  # points for each filled.
        + points_preferred * pulp.lpSum(assign_vars[w][shift] for w in workers for shift in empty_shifts if shift in worker_prefers[w])  # +points for preferred.
        + points_spacing * pulp.lpSum(spacing_var[w][pair] for w in workers for pair in bad_spacing_pairs)  # points for each bad spacing pair assigned to same w (both 1).
        + points_24hr * pulp.lpSum(twenty_four_vars[w][pair] for w in workers for pair in twenty_four_hour_shift_pairs) # points for each 24-hour shift
        )

        # Constraints (rules).
        # 1. At most one worker per shift (allow 0 to fill as many as possible).
        for shift in empty_shifts:
            prob += pulp.lpSum(assign_vars[w][shift] for w in workers) <= 1  # 0 or 1.
        
        # 2. Shifts per worker match range.
        for w in workers:
            min_shifts, max_shifts = next(worker["shifts_to_fill"] for worker in workers_list if worker["name"] == w)
            prob += pulp.lpSum(assign_vars[w][shift] for shift in empty_shifts) <= max_shifts  # At most max.
            prob += pulp.lpSum(assign_vars[w][shift] for shift in empty_shifts) >= min_shifts  # At least min.
        
        # 3. No night to day (hard rule).
        for pair in bad_night_to_day_pairs:
            night, day = pair  # Unpack the pair.
            for w in workers:
                prob += assign_vars[w][night] + assign_vars[w][day] <= 1  # Can't be 2 (both).        

        # 4. No adjacent nights (hard rule).
        for pair in bad_adjacent_nights_pairs:
            current, next_shift = pair  # Unpack the pair.
            for w in workers:
                prob += assign_vars[w][current] + assign_vars[w][next_shift] <= 1  # Can't be 2 (both).

        # 5. No adjacent days (hard rule).
        for pair in bad_adjacent_days_pairs:
            current, next_shift = pair  # Unpack the pair.
            for w in workers:
                prob += assign_vars[w][current] + assign_vars[w][next_shift] <= 1  # Can't be 2 (both).

        # 6. Cap limit for 24-hour shifts (Day X and Night X same day).
        # Use twenty_four_hour_shift_pairs – but filter for empty (PuLP can affect).
        # Hard limit: Sum of 24-hour for w <= max_24hr.
        # Soft: -2 points for each 24-hour.

        # Link: For each pair, twenty_four_vars[w][pair] == assign_vars[w][day] + assign_vars[w][night] == 2 (but since binary, if both 1, =1).
        for pair in twenty_four_hour_shift_pairs:
            day, night = pair
            for w in workers:
                prob += twenty_four_vars[w][pair] >= assign_vars[w][day] + assign_vars[w][night] - 1  # If both 1, >=1, so =1.
                prob += twenty_four_vars[w][pair] <= assign_vars[w][day]  # If not day, <=0.
                prob += twenty_four_vars[w][pair] <= assign_vars[w][night]  # If not night, <=0.

        # Hard limit: Sum twenty_four_vars[w][pair] for all pairs <= max_24hr[w].
        for w in workers:
            prob += pulp.lpSum(twenty_four_vars[w][pair] for pair in twenty_four_hour_shift_pairs) <= max_24hr[w]  # Max.

        # 7: Max weekend shifts
        for w in workers:
            prob += pulp.lpSum(assign_vars[w][shift] for shift in empty_weekend_shifts) <= max_weekends[w]  # <= max.

        # 8: Bad spacing pairs
        # Rules to set the extra box right (similar to 24-hour).
        for pair in bad_spacing_pairs:  # Loop bad pairs.
            s1, s2 = pair  # Unpack (s1 = first, s2 = second).
            for w in workers:  # For each worker.
                prob += spacing_var[w][pair] >= assign_vars[w][s1] + assign_vars[w][s2] - 1  # If both 1, >=1 (must be 1).
                prob += spacing_var[w][pair] <= assign_vars[w][s1]  # If not s1, <=0 (must be 0).
                prob += spacing_var[w][pair] <= assign_vars[w][s2]  # If not s2, <=0.

        ## Fence for "cannot work" – don't assign worker to "NO" shifts.
        for w in workers:  # Loop each worker name.
            for f in worker_cannot[w]:  # Loop each "NO" shift f for w.
                if f in empty_shifts:  # If f is empty (PuLP can assign it).
                    prob += assign_vars[w][f] <= 0  # Must be 0 (no assign).


        # Make the map to CBC.exe – like telling PuLP where the tool is
        if hasattr(sys, '_MEIPASS'):  # Check if we're in .exe mode (a special box PyInstaller adds)
            cbc_path = os.path.join(sys._MEIPASS, 'cbc.exe')  # Glue temp spot + file name
        else:  # Normal .py mode
            cbc_path = 'cbc.exe'  # Use the one in your folder
        print(f"Using CBC path: {cbc_path}")  # Note to check in console

        # Solve with messages and time limit.
        print("Starting PuLP solve – time limit 60 seconds...")
        status = prob.solve(pulp.COIN_CMD(msg=1, timeLimit=60, path=cbc_path))  # msg=1 shows progress, timeLimit=60 seconds.
        #status = prob.solve(pulp.PULP_CBC_CMD(msg=1, timeLimit=60))if cbc_path not defined
        if pulp.LpStatus[status] == "Optimal":
            print("Found best solution in time!")
        elif pulp.LpStatus[status] == "Not Solved":
            print("Timed out – showing best found so far.")
        else:
            print("Other issue:", pulp.LpStatus[status])

        print(f"Solve finished! Status: {pulp.LpStatus[status]}")  # Shows if it worked (e.g., "Optimal")

        # Add assignments (same as before, even if timed out).
        for w in workers:
            for shift in empty_shifts:
                if assign_vars[w][shift].value() == 1:
                    assignments[shift] = w
        print("Assignments done!")

        # Print the total points (the score from the goal).
        total_points = prob.objective.value()  # Get the number from PuLP's goal box.
        print("Total points for this rota:", total_points)  # Show it.

    # Show final assignments.
    print("Final Rota:")
    for shift, worker in assignments.items():
        print(f"{shift}: {worker if worker else 'Unassigned'}")

    # Count and print the summary at the end (how many preferences, 24-hour, bad spacing).
    preferences_count = 0  # Start count at 0 – like an empty basket for good things.
    twenty_four_count = 0  # For 24-hour.
    bad_spacing_count = 0  # For bad close.

    # Get worker prefers and cannot (boxes from earlier).
    worker_prefers = {w["name"]: w["prefers"] for w in workers_list}  # Worker: list of preferred.

    # Loop all assignments to count.
    for shift, worker in assignments.items():  # Loop each shift and who (worker or None).
        if worker is None:  # Skip empty.
            continue

        # Count preferences: If shift in worker's prefers.
        if shift in worker_prefers[worker]:  # If preferred.
            preferences_count += 1  # Add 1.

    # For 24-hour and spacing, reuse grouping.
    worker_shifts = {}  # Box worker: list of their shifts.
    for shift, worker in assignments.items():
        if worker is not None:
            if worker not in worker_shifts:
                worker_shifts[worker] = []
            worker_shifts[worker].append(shift)

    for worker, shifts in worker_shifts.items():  # Loop workers and their shifts.
        sorted_shifts = sorted(shifts, key=lambda s: int(s.split(" ")[1]))  # Sort by day.
        for i in range(len(sorted_shifts) - 1):  # Loop pairs.
            current = sorted_shifts[i]
            next_shift = sorted_shifts[i+1]
            current_day = int(current.split(" ")[1])
            next_day = int(next_shift.split(" ")[1])
            if current_day == next_day:  # Same day – 24-hour.
                twenty_four_count += 1  # Add 1.
            if next_day - current_day < 5:  # Close – bad spacing.
                bad_spacing_count += 1  # Add 1.

    # Print the counts.
    print("Summary:")
    print("Number of preferred shifts assigned:", preferences_count)
    print("Number of 24-hour shifts:", twenty_four_count)
    print(f"Number of bad spacing pairs (<{spacing_days_threshold} days apart, not shifts):", bad_spacing_count)
    
    # Create popup to display results
    popup = Toplevel(root)
    popup.title("Rota Results")

    # Create a Text widget instead of using Labels in a grid
    # Text widget is like a text editor - you can select and copy from it
    text_widget = Text(popup, wrap="none", width=60, height=30)
    text_widget.pack(side="left", fill="both", expand=True)

    # Create scrollbar for the Text widget
    scrollbar = Scrollbar(popup, orient="vertical", command=text_widget.yview)
    scrollbar.pack(side="right", fill="y")
    text_widget.config(yscrollcommand=scrollbar.set)

    def on_popup_mousewheel(event):
        text_widget.yview_scroll(int(-1 * (event.delta / 120)), "units")
    
    popup.bind("<MouseWheel>", on_popup_mousewheel)
    text_widget.bind("<MouseWheel>", on_popup_mousewheel)

    # Now we INSERT text into the Text widget instead of creating Labels
    # Start building the text to display
    text_widget.insert("end", "Final Rota:\n", "title")  # "end" means add to the end
    text_widget.insert("end", "=" * 60 + "\n\n", "separator")  # Line of equals signs

    # Add headers
    text_widget.insert("end", f"{'Day':<7}{'Day Shift':<25}{'Night Shift':<15}\n", "header")
    text_widget.insert("end", "-" * 60 + "\n", "separator")

    # Add assignments
    for day in days_list:
        day_shift = f"Day {day}"
        night_shift = f"Night {day}"
        day_worker = assignments.get(day_shift, "No shift")
        night_worker = assignments.get(night_shift, "No shift")
        
        # Format the line nicely with spacing
        day_text = str(day)
        day_worker_text = day_worker if day_worker else "Unassigned"
        night_worker_text = night_worker if night_worker else "Unassigned"
        
        text_widget.insert("end", f"{day_text:<5}{day_worker_text:<15}{night_worker_text:<15}\n")

    # Add summary section
    text_widget.insert("end", "\n" + "=" * 60 + "\n", "separator")
    text_widget.insert("end", "Summary:\n", "title")
    text_widget.insert("end", "-" * 60 + "\n", "separator")
    text_widget.insert("end", f"Number of preferred shifts assigned: {preferences_count}\n")
    text_widget.insert("end", f"Number of 24-hour shifts: {twenty_four_count}\n")
    text_widget.insert("end", f"Number of bad spacing pairs (<5 days apart): {bad_spacing_count}\n")

    # Make some text bold/bigger (optional styling)
    text_widget.tag_config("title", font=("Arial", 12, "bold"))
    text_widget.tag_config("header", font=("Arial", 10, "bold"))
    text_widget.tag_config("separator", foreground="gray")

    # Make the text read-only (so users can't accidentally edit it)
    # But they CAN still select and copy!
    text_widget.config(state="disabled")

# Add worker button.
Button(add_worker_button_frame, text="Add Worker", command=add_worker_row, width=10, pady=2).pack() 

# Create rota button.
Button(root, text="Create Rota", command=create_rota).pack()  # Button to create a rota.

# Frame for settings, saving and loading buttons.
settings_frame = Frame(root)
settings_frame.pack()

Button(settings_frame, text="PuLP Settings", width=18, command=pulp_settings).pack(side=LEFT, padx=5)
Button(settings_frame, text="Save Preferences", width=18, command=save_preferences).pack(side=LEFT, padx=5)
Button(settings_frame, text="Load", width=18, command=load_preferences).pack(side=LEFT, padx=5)

# Error label (same).
error_label = Label(root, text="")  # For errors.
error_label.pack()

root.mainloop()  # Start the window – like "go!"
