import sys # Easy way for Python to restart
import pulp  # Bring in the PuLP toolbox.
import time # Used to show how much time it takes to solve the rota
import calendar  # Bring in the calendar toolbox for month days.
import os # Bring in interaction with Windows/Apple/Linux
import json # Bring in functionality of saving/loading javascript object notation - data-interchange format
from tkinter import *  # Bring in the Tkinter toolbox for the window (GUI). Does not import modules within Tkinter, only functions like Button, Label etc. From ... import syntax brings brings namespace to THIS current namespace
from tkinter import filedialog # A module (.py file) within tkinter, so it has to be imported separately
import openpyxl # Allows to use .xlsx files
from selection_popups import prefer_count, cannot_count, prefer_unit_count, manual_count # bring popup_select_shifts functions from a another file
from solver import solve_rota # bring PuLP solver from another file
from date_settings import save_year_confirm, save_month_confirm, save_holidays_confirm
from pulp_settings import pulp_settings
from save_load import save_preferences as save_preferences_file, load_preferences as load_preferences_file, load_xlsx_preferences as load_xlsx_preferences_file
import threading # Allows to run the Tkinter GUI while solving rota (the solving part runs separate)
import psutil      # Lets us find child processes by parent. Later the subprocesses can be killed, important to shut down cbc.exe midway.
import subprocess # When closing main app, closes all children windows (like CBC.exe or Windows console)

# --------------------------------------------------------------------
# App plan:
# Part I: Tkinter GUI
# Part II: PuLP Solve
# --------------------------------------------------------------------

# --------------------------------------------------------------------
# Tkinter GUI
# --------------------------------------------------------------------

# Global variables: IMPORTANT!!!
year = None
month = None
holiday_days = [] # Empty variable, so "Make shifts" works even if Holidays saved nothing
shifts_list = []
workers_list = []
units_list = [] # Empty list to store units: "Cardiology", "Internal Medicine - Endocrinology" etc.
selected_cannot_days = {}  # Changed to dict to store per row_num
selected_prefer_days = {}  # Dict to store preferred days per row_num
selected_units = {}        # Dict to store preferred units per row_num
selected_manual_days = {}  # Dict to store manual days per row_num

# Global variables: constant
day_names = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday", "Sunday"]  # Group of day names.

# Settings for PuLP: points, hard rules; Other settings: shift making
points_filled = 100
points_preferred = 1
points_preferred_unit = 3
points_spacing = -2
spacing_days_threshold = 4  # How many days apart triggers the penalty
points_24hr = -3
enforce_no_adj_nights = True   # Night → Night next day
enforce_no_adj_days = True   # Day → Day next day hard rule
include_weekday_days = False   # False = default behaviour (skip Mon-Fri day shifts when making shifts)

# -------------------------------------------------------
# The main window (like the car's dashboard).
# -------------------------------------------------------

root = Tk()  # Make the window box.
#root.geometry("975x450")  # Set predetermined window size (width x height)
root.title("Hospital Shift Manager 'Riaukapp'")  # Name on top.

# Top frame designed for LabelFrame and Utilities frame
top_frame = Frame(root)
top_frame.pack(padx=8, pady=(8, 2), fill="x", anchor="w")

# LabelFrame groups year, month, holidays under one visible "Date Settings" border
date_frame = LabelFrame(top_frame, text="Date Settings", padx=6, pady=4)
date_frame.pack(side=LEFT, fill="both", expand=True, anchor="nw")

utilities_lf = LabelFrame(top_frame, text="Utilities", padx=10, pady=10)
utilities_lf.pack(side=LEFT, fill="both", expand=True, anchor="nw", padx=(8, 0))

Label(date_frame, text="Year (e.g., 2026):").pack(anchor="w")

# Frame for year row – like a shelf for horizontal.
year_frame = Frame(date_frame)  # Parent is now date_frame, not root
year_frame.pack(anchor="w")

Label(date_frame, text="Month (1-12, e.g., 1 for January):").pack(anchor="w")

# Type box inside frame.
year_entry = Entry(year_frame)  # Type box in frame.
year_entry.pack(side=LEFT)  # Next to label.

# Function for save year (like button press).
def save_year():
    global year, month, num_days, starting_weekday, days_list

    save_year_inputs = {
        "give_year_entry": year_entry,
        "give_error_label": error_label,
        "give_current_year_label": current_year_label,
        "give_year": year,
    }
    
    result = save_year_confirm(save_year_inputs)
    if result is not None:
        year = result
        # Reset month data — it was calculated for the old year
        if month is not None:
            month = None
            num_days = None
            starting_weekday = None
            days_list.clear()
            holiday_days.clear()
            current_month_label.config(text="Current Month: None")
            holidays_label.config(text="Holiday List: None")
            error_label.config(text="Year changed, so month and holidays have been reset!")

# Button inside frame.
Button(year_frame, text="Save", command=save_year).pack(side=LEFT)  # Button next.

# Show current year.
current_year_label = Label(year_frame, text="Current Year: None")  # Show label.
current_year_label.pack(side=LEFT)

month_frame = Frame(date_frame)  # Parent is date_frame, not root
month_frame.pack(anchor="w")

# Label and type box for month.
month_entry = Entry(month_frame)  # Type box.
month_entry.pack(side=LEFT)

def save_month():
    global year, month, num_days, starting_weekday, days_list

    if workers_list:
        error_label.config(text="Error: Cannot change months after workers have been added. Remove all workers first.")
        return  

    save_month_inputs = {
        "give_month_entry": month_entry,
        "give_error_label": error_label,
        "give_current_month_label": current_month_label,
        "give_month": month,
        "give_year": year,
    }

    # Handles None safely, for example if year was not set, the function would return None, but this code would try to unpack the tuple (month, num_days etc.), but it can't because it's None!
    result = save_month_confirm(save_month_inputs)
    if result is not None:
        month, num_days, starting_weekday, days_list = result
        if holiday_days:
            holiday_days.clear()
            holidays_label.config(text="Holiday List: None")
            error_label.config(text="Month changed, so holidays were reset!")

Button(month_frame, text="Save", command=save_month).pack(side=LEFT)  # Button, click runs save_month.

current_month_label = Label(month_frame, text="Current Month: None")
current_month_label.pack(side=LEFT)

Label(date_frame, text="Public holidays, comma-separated (e.g., 24,25,26) or leave blank:").pack(anchor="w")

holiday_frame = Frame(date_frame)  # Parent is date_frame, not root
holiday_frame.pack(anchor="w")

# Label and type box for holidays.
holiday_entry = Entry(holiday_frame)  # Type box.
holiday_entry.pack(side=LEFT)

def save_holidays():  # Function for the button.
    global holiday_days  # Use the holiday_days box outside.

    save_holidays_inputs = {
        "give_year": year,
        "give_month": month,
        "give_holiday_entry": holiday_entry,
        "give_holidays_label": holidays_label,
        "give_error_label": error_label,
    }

    result = save_holidays_confirm(save_holidays_inputs)
    if result is not None:
        holiday_days = result
    
Button(holiday_frame, text="Save", command=save_holidays).pack(side=LEFT)  # Button, click runs save_holidays.

holidays_label = Label(holiday_frame, text="Holiday List: None")
holidays_label.pack(side=RIGHT)

# Separate LabelFrame for Units, with groove relief
units_lf = LabelFrame(root, text="Units", relief="groove", padx=6, pady=4)
units_lf.pack(padx=8, pady=(0, 4), fill="x", anchor="w")

Label(units_lf, text="Comma-separated (e.g. Internal Medicine,Cardiology):").pack(anchor="w")

units_frame = Frame(units_lf)  # Parent is units_lf, not root
units_frame.pack(anchor="w")

units_entry = Entry(units_frame, width=40)
units_entry.pack(side=LEFT)



def save_units():
    global units_list

    if workers_list:
        error_label.config(text="Error: Cannot change units after workers have been added. Remove all workers first.")
        return

    units_input = units_entry.get().strip()

    # Split by comma and clean up whitespace
    units_list = [u.strip() for u in units_input.split(",") if u.strip()] # give me u.strip(), for each u in the split list, but only if u.strip() is not empty".
    
    if not units_list:
        current_units_label.config(text="Current Units: None")
        error_label.config(text="Please enter at least one unit.")
        return
    
    current_units_label.config(text=f"Current Units: {', '.join(units_list)}")
    error_label.config(text=f"{len(units_list)} unit(s) saved.")

Button(units_frame, text="Save", command=save_units).pack(side=LEFT)

current_units_label = Label(units_frame, text="Current Units: None")
current_units_label.pack(side=LEFT)

# =======================================
# Popup functions, found in separate file
# =======================================

def show_prefer_popup(row_num):
    """
    Opens a popup window where the user can choose which shifts this worker
    would LIKE to work (preferred days/shifts).

    What it does, step by step:
    1. Creates a new small window with title "Select Days Prefer Work"
    2. Lists every day of the month + checkboxes for Day and Night shifts
    3. Pre-ticks boxes that were already selected before (if the user opened it again)
    4. Includes "Check All Day" and "Check All Night" boxes for fast selection
    5. When user clicks "Save Selection":
       - Saves chosen preferred shifts into selected_prefer_days[row_num]
       - Changes button text to show how many were selected (e.g. "Select (4)")
       - Shows message like "Worker prefers to work on these shifts: ..."
       - Closes the popup

    - Similar function for the other popups
    """

    prefer_popup_inputs = {
        "give_root": root,
        "give_days_list": days_list,
        "give_starting_weekday": starting_weekday,
        "give_worker_rows": worker_rows,
        "give_row_num": row_num,
        "give_selected_prefer_days": selected_prefer_days,   
        "give_error_label": error_label,                     
    }

    prefer_count(prefer_popup_inputs)


def show_cannot_popup(row_num):

    cannot_popup_inputs = {
        "give_root": root,
        "give_days_list": days_list,
        "give_starting_weekday": starting_weekday,
        "give_worker_rows": worker_rows,
        "give_row_num": row_num,
        "give_selected_cannot_days": selected_cannot_days,   
        "give_error_label": error_label,                
    }

    cannot_count(cannot_popup_inputs)

def show_prefer_unit_popup(row_num):

    prefer_unit_popup_inputs = {
        "give_root": root,
        "give_days_list": days_list,
        "give_starting_weekday": starting_weekday,
        "give_worker_rows": worker_rows,
        "give_row_num": row_num,
        "give_selected_manual_days": selected_manual_days,   
        "give_error_label": error_label,
        "give_units_list": units_list,
        "give_selected_units": selected_units,    
    }
    
    prefer_unit_count(prefer_unit_popup_inputs)

def show_manual_popup(row_num):

    manual_popup_inputs = {
        "give_root": root,
        "give_days_list": days_list,
        "give_starting_weekday": starting_weekday,
        "give_worker_rows": worker_rows,
        "give_row_num": row_num,
        "give_selected_manual_days": selected_manual_days,   
        "give_error_label": error_label,
        "give_units_list": units_list,    
    }
    
    manual_count(manual_popup_inputs)




# Button to add row.

# Frame to hold the scrollable worker area (like a shelf for the canvas)
worker_container = Frame(root)
#worker_container.pack(fill="both", expand=True)  # Fill the space, can grow
worker_container.pack()

# Create the canvas (drawing board)
global worker_canvas, worker_scrollbar, worker_inner_frame
worker_canvas = Canvas(worker_container, width=833, borderwidth=2, relief="groove")  # Height=200 pixels – change if you want taller/shorter
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

def make_shifts():  # Function for making shifts list.
  
    # Check if year AND month are defined
    if year == None or month == None:
        error_label.config(text="Erorr: Please select year and month first!")
        return

    # Check if units are defined
    if not units_list:
        error_label.config(text="Error: Please define units first!")
        return

    shifts_list.clear()  # Empty list for shifts.
    shift_types = ["Day", "Night"]  # Group of types.
    
    # Loop through each unit
    for unit in units_list:
        for day in days_list:  # Loop each day.
            weekday = (starting_weekday + (day - 1)) % 7  # Calculate day name ( %7 like clock wrap).
            tags = [day_names[weekday]]  # Add day name tag.
            if weekday in [5, 6]:  # Sat/Sun – weekend.
                tags.append("Weekend")  # Add tag.
            if day in holiday_days:  # If holiday.
                tags.append("Public holiday")  # Add tag.
            for shift_type in shift_types:  # Loop Day then Night.
                # Check if this is a Monday-Friday Day shift (and not a holiday) – if both true, exclude (skip)
                if (not include_weekday_days) and weekday in [0, 1, 2, 3, 4] and shift_type == "Day" and day not in holiday_days:
                    continue  # Skip Mon-Fri day shifts only when the checkbox is OFF
                shift_name = f"{shift_type} {day} {unit}"  # Make name: ex. Cardiology Day 1, Internal Medicine Night 2...
                shift_dict = {  # Make the shift box.
                    "name": shift_name,
                    "type": shift_type,
                    "tags": tags,
                    "unit": unit,
                    "assigned_worker": None
                }
                shifts_list.append(shift_dict)  # Add to list.
        
        error_label.config(text=f"Shifts made: {len(shifts_list)} across {len(units_list)} unit(s)") # Update the label with count.
        
        # Show the shifts list nicely in terminal, for debugging
        print("Your shifts list (with tags, types):")
        for shift in shifts_list:  # Loop to print each one.
            print(f"Shift: {shift['name']}, Tags: {shift['tags']}, Type: {shift['type']}, Assigned: {shift['assigned_worker']}")

# Box to show the shifts list.
#shifts_label = Label(root, text="Shifts made: None")  # Start None.
#shifts_label.pack()  # Put on window.

# Frame for workers – like a big shelf for the table.
workers_frame = Frame(worker_inner_frame)  # Pack to inner_frame
workers_frame.pack(fill="x")  # Fill horizontal, stack vertical

'''
The layout of all worker containers:

worker_container        → positions canvas + scrollbar side by side (pack LEFT / RIGHT)
  └── worker_canvas     → provides the scrollable viewport
        └── worker_inner_frame  → the actual scrollable surface (grows as rows are added)
              └── workers_frame → owns the grid layout (headers + data rows)

The first three layers are the standard Tkinter scrollable area pattern — you need all three
because a Canvas can't directly manage a grid of widgets; you embed a Frame inside it as a workaround.

Why workers_frame as a 4th layer?
It's because worker_inner_frame and workers_frame use different layout managers:
- worker_inner_frame uses pack — it stacks things vertically as rows are added
- workers_frame uses grid — it arranges the header labels and data cells in columns
You cannot mix pack and grid in the same parent widget. 

In other words:
worker_inner_frame uses pack to arrange its children → workers_frame is packed into it ✓
workers_frame uses grid to arrange its own children → all the labels and entries are gridded into it ✓

So workers_frame exists purely to isolate the grid layout in its own container, 
keeping pack and grid separated into different parent widgets              
'''
              
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
Label(workers_frame, text="Prefer Unit", width=10, padx=2, pady=2).grid(row=0, column=6, sticky="ew", padx=2, pady=2)
Label(workers_frame, text="Manual Shifts", width=10, padx=2, pady=2).grid(row=0, column=7, sticky="ew", padx=2, pady=2)
Label(workers_frame, text="Save", width=10, padx=2, pady=2).grid(row=0, column=8, sticky="ew", padx=2, pady=2)
Label(workers_frame, text="Delete", width=10, padx=2, pady=2).grid(row=0, column=9, sticky="ew", padx=2, pady=2)

# List to hold worker rows (for delete).
worker_rows = []  # Empty list to store widget references for each row.
worker_row_number = 1  # Start row 1 (after 0 headers)

# The worker_row_number is never reset to 1 when deleting worker rows
# because it serves as a unique ID and a key in dictionaries like selected_cannot_days etc.

def add_worker_row():  # Function for "Add Worker" button.
    global worker_row_number, workers_list
    
    if year is None or month is None:
        error_label.config(text="Error: Please set year and month before adding workers.")
        return

    if not units_list:
        error_label.config(text="Error: Please define units before adding workers.")
        return
    
    row_num = worker_row_number  # Capture the current row number for this row

    # Column 0: Name box - placed directly in workers_frame.
    name_entry = Entry(workers_frame, width=10)  # Type box.
    name_entry.grid(row=row_num, column=0, sticky="ew", padx=2, pady=2)

    # Column 1: Shift range box.
    range_entry = Entry(workers_frame, width=10)
    range_entry.grid(row=row_num, column=1, sticky="ew", padx=2, pady=2)

    # Column 2: Cannot work button.
    cannot_button = Button(workers_frame, text="Select", width=10, command=lambda: show_cannot_popup(row_num))
    cannot_button.grid(row=row_num, column=2, sticky="ew", padx=2, pady=2)

    # Column 3: Prefer button.
    prefer_button = Button(workers_frame, text="Select", width=10, command=lambda: show_prefer_popup(row_num))
    prefer_button.grid(row=row_num, column=3, sticky="ew", padx=2, pady=2)

    # Column 4: Max weekends box.
    max_weekends_entry = Entry(workers_frame, width=10)
    max_weekends_entry.grid(row=row_num, column=4, sticky="ew", padx=2, pady=2)

    # Column 5: Max 24-hour box.
    max_24hr_entry = Entry(workers_frame, width=10)
    max_24hr_entry.grid(row=row_num, column=5, sticky="ew", padx=2, pady=2)

    # Column 6: Prefer_unit_button.
    prefer_unit_button = Button(workers_frame, text="Select", width=10, command=lambda: show_prefer_unit_popup(row_num))
    prefer_unit_button.grid(row=row_num, column=6, sticky="ew", padx=2, pady=2)

    # Column 7: Manual shifts button.
    manual_button = Button(workers_frame, text="Select", width=10, command=lambda: show_manual_popup(row_num))
    manual_button.grid(row=row_num, column=7, sticky="ew", padx=2, pady=2)

    # Column 8: Save Worker button.
    save_button = Button(workers_frame, text="Save", width=10, command=lambda: save_worker(row_num))
    save_button.grid(row=row_num, column=8, sticky="ew", padx=2, pady=2)

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
        'prefer_unit_button': prefer_unit_button,
        'manual_button': manual_button,
        'save_button': save_button,
        'delete_button': delete_button,
        'row_num': row_num
    }
    worker_rows.append(row_widgets)
    print(f"Current row: {worker_row_number}") #Debugging
    worker_row_number += 1  # Add 1. 
    
    def save_worker(row_num):
        """
        Validates and saves (or updates) a worker from the GUI row identified by row_num.

        Reads input from the Entry widgets captured in the enclosing add_worker_row() scope.
        Validates that name is not blank, shift range is in "min-max" format with valid numbers,
        and that max_weekends / max_24hr are non-negative integers (default 100 if left blank).
        Also reads cannot_work days, preferred days, and preferred units from their global dicts.

        If validation passes:
        - If a worker with this row_num already exists in workers_list, updates their data.
        - Otherwise, creates a new worker dict and appends it to workers_list.

        Args:
            row_num (int): The unique ID of this worker row, used as a key in all
                        per-row dictionaries (selected_cannot_days, etc.).
        """

        name = name_entry.get().strip() # Get the name
        range_input = range_entry.get().strip()  # Get shift range.
        max_weekends_input = max_weekends_entry.get().strip()  # Get max weekends.
        max_24hr_input = max_24hr_entry.get().strip()  # Get max 24-hour.
        
        # Cannot work and other from selected days.
        cannot_work = selected_cannot_days.get(row_num, [])  # Get the list for this row_num.
        prefer = selected_prefer_days.get(row_num, [])  # Get the list for this row_num.
        prefer_units = selected_units.get(row_num, []) # Get the list for this row_num for prefer_units

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
            existing_worker["prefer_units"] = prefer_units
            existing_worker["max_weekends"] = max_weekends
            existing_worker["max_24hr"] = max_24hr
            # Show message if succesful
            message = f"Worker '{name}' updated!"
            error_label.config(text=message)  # Show success.
        else:  # If not found (new worker)
            # Create new worker dictionary
            worker_dict = {
                "name": name,
                "shifts_to_fill": [min_shifts, max_shifts],
                "cannot_work": cannot_work,
                "prefers": prefer,
                "prefer_units": prefer_units,
                "max_weekends": max_weekends,
                "max_24hr": max_24hr,
                "worker_row_number": row_num
            }
            workers_list.append(worker_dict)  # Add new worker.
            error_label.config(text=f"Worker '{name}' saved!")  # Show success.

        # Print for debugging
        print("Your full worker list:")
        for worker in workers_list:
            print(f"Name: {worker['name']}, shifts: {worker['shifts_to_fill']}, Cannot: {worker['cannot_work']}, Prefers: {worker['prefers']}, Max weekends: {worker['max_weekends']}, Max 24hr: {worker['max_24hr']}, Prefers units: {worker['prefer_units']} Row number: {worker['worker_row_number']}")

        '''
        Note on dictionaries: both worker_dict and separate selected_cannot_days + other dictionaries store similar data.
        The solver (solver.py) reads cannot_work, prefers, and prefer_units directly from the worker dict in workers_list. It never touches selected_cannot_days, selected_prefer_days, or selected_units at all.
        So these three fields must be stored in the worker dict — the solver depends on them being there. Without them, the solver would have no cannot/prefer information and the rota would be generated ignoring all those preferences.
        selected_manual_days works differently because manual shifts are pre-assigned to shifts_list before the solver even runs (via assign_all_manual_shifts() on line 1442), so the solver never needs to look them up from the worker dict.
        In summary:
        cannot_work, prefers, prefer_units → must be in the worker dict (solver reads them)
        manual days → don't need to be in the worker dict (handled before the solver runs)
        '''

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

def assign_all_manual_shifts():
    """
    Loops through every worker and assigns their manually selected shifts.
    
    Called automatically inside create_rota(), AFTER make_shifts().
    Does NOT change shift ranges in workers_list or update GUI widgets.
    """
    for worker in workers_list:
        row_num = worker["worker_row_number"]
        name = worker["name"]
        
        # Get this worker's manually selected shifts
        # .get() returns [] as default if row_num not found - avoids a crash
        manual_shifts = selected_manual_days.get(row_num, [])
        
        for shift_name in manual_shifts: # For example Day 2 Cardiology from Manual shifts ["Day 2 Cardiology", "Night 3 Internal Medicine"]
            # Search through shifts_list to find the matching shift
            for shift in shifts_list:
                if shift["name"] == shift_name:
                    # Only assign if not already taken
                    if shift["assigned_worker"] is None:
                        shift["assigned_worker"] = name
                    else:
                        print(f"Warning: {shift_name} already assigned to {shift['assigned_worker']}, skipping.")
                    break  # Found the shift - no need to keep searching

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

    # Clean up the dictionaries when deleting
    if row_num in selected_cannot_days:
        del selected_cannot_days[row_num]
    if row_num in selected_prefer_days:
        del selected_prefer_days[row_num]
    if row_num in selected_units:
        del selected_units[row_num]
    if row_num in selected_manual_days:
        del selected_manual_days[row_num]

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
        '''
        The loop iterates over every key/value pair and destroys each one. The two conditions:
        key != 'row_num' — skips the 'row_num' entry because it's a plain integer, not a widget. Calling .destroy() on an integer would crash.
        widget.winfo_exists() — checks the widget still exists in Tkinter before destroying it. This guards against a case where a widget may have already been destroyed (e.g. if its parent was destroyed first, children go with it).
        So together the two checks ensure: only destroy things that are actual live widgets.
        '''

    # Remove the worker from workers_list if it exists
    workers_list = [worker for worker in workers_list if worker.get("worker_row_number") != row_num]
    error_label.config(text=f"Worker '{worker_name}' deleted and all their shifts unassigned!")

def open_pulp_settings():
    """
    Gathers current PuLP settings into a dict and opens the settings popup.
    Updates global settings variables when the user saves.
    """

    settings_inputs = {
        "points_filled":          points_filled,
        "points_preferred":       points_preferred,
        "points_preferred_unit":  points_preferred_unit,
        "points_spacing":         points_spacing,
        "spacing_days_threshold": spacing_days_threshold,
        "points_24hr":            points_24hr,
        "enforce_no_adj_days":    enforce_no_adj_days,
        "enforce_no_adj_nights":  enforce_no_adj_nights,
        "include_weekday_days":   include_weekday_days,
    }

    def apply_new_settings(new_settings):
        global points_filled, points_preferred, points_preferred_unit, points_spacing
        global spacing_days_threshold, points_24hr, enforce_no_adj_days, enforce_no_adj_nights, include_weekday_days
        points_filled          = new_settings["points_filled"]
        points_preferred       = new_settings["points_preferred"]
        points_preferred_unit  = new_settings["points_preferred_unit"]
        points_spacing         = new_settings["points_spacing"]
        spacing_days_threshold = new_settings["spacing_days_threshold"]
        points_24hr            = new_settings["points_24hr"]
        enforce_no_adj_days    = new_settings["enforce_no_adj_days"]
        enforce_no_adj_nights  = new_settings["enforce_no_adj_nights"]
        include_weekday_days   = new_settings["include_weekday_days"]

    pulp_settings(root, settings_inputs, error_label, apply_new_settings)

def save_preferences():
    
    save_preferences_file(error_label, {
        "year": year,
        "month": month,
        "units_list": units_list,
        "workers_list": workers_list,
        "holiday_days": holiday_days,
        "shifts_list": shifts_list,
        "selected_cannot_days": selected_cannot_days,
        "selected_prefer_days": selected_prefer_days,
        "selected_units": selected_units,
        "selected_manual_days": selected_manual_days,
    })

def load_preferences():
    def set_worker_row_number(n):
        global worker_row_number
        worker_row_number = n

    def set_year(y):
        global year
        year = y

    def set_month(m):
        global month
        month = m

    load_preferences_file(
        widgets={
            "error_label": error_label,
            "current_year_label": current_year_label,
            "year_entry": year_entry,
            "month_entry": month_entry,
            "current_month_label": current_month_label,
            "holiday_entry": holiday_entry,
            "holidays_label": holidays_label,
            "units_entry": units_entry,
            "current_units_label": current_units_label,
        },
        state={
            "workers_list": workers_list,
            "worker_rows": worker_rows,
            "selected_cannot_days": selected_cannot_days,
            "selected_prefer_days": selected_prefer_days,
            "selected_units": selected_units,
            "selected_manual_days": selected_manual_days,
            "holiday_days": holiday_days,
            "units_list": units_list,
            "shifts_list": shifts_list,
        },
        callbacks={
            "save_month": save_month,
            "add_worker_row": add_worker_row,
            "set_worker_row_number": set_worker_row_number,
            "set_year": set_year,
            "set_month": set_month,
        }
    )


def load_xlsx_preferences():
    if year is None or month is None:
        error_label.config(text="Please set year and month before loading an xlsx file.")
        return
    
    def set_worker_row_number(n):
        global worker_row_number
        worker_row_number = n

    result = load_xlsx_preferences_file(
        widgets={
            "error_label": error_label,
            "units_entry": units_entry,
            "current_units_label": current_units_label,
            "root": root,
        },
        state={
            "workers_list": workers_list,
            "worker_rows": worker_rows,
            "selected_cannot_days": selected_cannot_days,
            "selected_prefer_days": selected_prefer_days,
            "selected_manual_days": selected_manual_days,
            "selected_units": selected_units,
            "units_list": units_list,
            "days_list": days_list,
        },
        callbacks={
            "add_worker_row": add_worker_row,
            "make_shifts": make_shifts,
            "set_worker_row_number": set_worker_row_number,
        }
    )
    if result:
        global worker_row_number
        worker_row_number = result["worker_row_number"]

def extract_day_from_shift_name(shift_name):
    """
    Extract day number from shift name.
    Examples: "Day 5 Cardiology" -> 5, "Night 12 Internal Medicine" -> 12
    """
    parts = shift_name.split()
    return int(parts[1])  # Day number is always the second part

def extract_unit_from_shift_name(shift_name):
    """
    Extract unit name from shift name.
    Examples: "Day 5 Cardiology" -> "Cardiology"
    """
    parts = shift_name.split()
    return " ".join(parts[2:])  # Everything after day number is the unit

# ----------------------------------------------------------------------------
# PuLP Solve: solver is in separate file
# ----------------------------------------------------------------------------

def set_solving_state(is_solving):
    """
    Disable or enable buttons while the solver is running.
    is_solving=True  → grey everything out (solver started)
    is_solving=False → restore everything (solver finished)
    
    WHY THIS APPROACH?
    The alternative would be using locks (threading.Lock) to protect
    shared data — but that's complex and can cause the GUI to freeze
    if a locked thread stalls. Disabling buttons is simpler and
    communicates clearly to the user what's happening.
    """
    # NORMAL state when not solving, DISABLED state when solving
    state = DISABLED if is_solving else NORMAL

    # Buttons that read or write shared data (shifts_list, workers_list etc.)
    buttons_to_control = [
        create_rota_button,     # Can't start a second solve
        add_worker_button,      # Can't add workers mid-solve
    ]
    for button in buttons_to_control:
        button.config(state=state)

    # Also disable all Save, Delete, Cannot, Prefer etc. buttons
    # on every worker row — these all touch workers_list
    for row_widgets in worker_rows:
        for key, widget in row_widgets.items():
            if key != 'row_num':
                widget.config(state=state)


def create_rota():
    """
    This is the create_rota function in main.py.
    It's much simpler - it just:
    1. Collects all the data from the GUI
    2. Packages it up
    3. Sends it to solver.py
    4. Gets the results back
    5. Shows the results to the user
    """
    # =======================================================================
    # Preparation: make shifts and assign workers with manual select shifts
    # =======================================================================
    make_shifts()  # Build fresh shifts list
    
    # Guard: if shifts_list is still empty, stop
    if not shifts_list:
        return  # make_shifts() already showed the error message
    
    assign_all_manual_shifts()  # Lock in manual assignments before solver runs

    # LOCK everything down immediately
    set_solving_state(True)
    # ========================================================================
    # PART A: Gather all the settings into a dictionary
    # ========================================================================
    # Think of this like packing a suitcase before a trip

    settings = {
        "points_filled": points_filled,
        "points_preferred": points_preferred,
        "points_spacing": points_spacing,
        "spacing_days_threshold": spacing_days_threshold,
        "points_24hr": points_24hr,
        "enforce_no_adj_nights": enforce_no_adj_nights,
        "enforce_no_adj_days": enforce_no_adj_days,
        "points_preferred_unit": points_preferred_unit                        
    }
    # -------------------------------------------------------
    # PART 1: The animated dots
    # -------------------------------------------------------
    # This list holds one item: the current dot count (0, 1, 2, or 3)
    # We use a list instead of a plain variable because of how Python
    # handles variables inside nested functions.
    #
    # WHY A LIST? If we wrote: dot_count = 0, then tried to do
    # dot_count += 1 inside animate(), Python would complain
    # "dot_count not defined" — because inner functions can READ
    # outer variables but can't REASSIGN them (without 'nonlocal').
    # A list gets around this: we never reassign dot_count itself,
    # we just change what's inside it: dot_count[0] += 1
    dot_count = [0]

    def animate():
        # dot_count[0] cycles: 0 → 1 → 2 → 3 → 0 → 1 → ...
        dot_count[0] = (dot_count[0] + 1) % 4  # % 4 means "wrap back to 0 after 3"
        dots = "." * dot_count[0]               # 0 dots, 1 dot, 2 dots, or 3 dots
        error_label.config(text=f"Solving rota{dots}")

        # root.after(milliseconds, function) is Tkinter's way of saying
        # "call this function again after X milliseconds — but only if
        # the solver is still running"
        # We store the "job ID" so we can cancel it later
        if not solver_done[0]:  # solver_done is explained below
            animate_job[0] = root.after(500, animate)  # 500ms = half a second

    # -------------------------------------------------------
    # PART 2: Flags to track if the solver has finished
    # -------------------------------------------------------
    # Same list trick as above — used to communicate between threads
    solver_done = [False]    # Becomes True when solver finishes
    animate_job = [None]     # Stores the root.after job ID so we can cancel it
    result_holder = [None]   # Will store (assignments, summary) when done

    # -------------------------------------------------------
    # PART 3: The solver function that runs in the background
    # -------------------------------------------------------
    def run_solver():
        # This whole function runs on a SEPARATE thread
        # so Tkinter stays free to animate
        try:
            assignments, summary = solve_rota(shifts_list, workers_list, units_list, settings)
            result_holder[0] = (assignments, summary)
            solver_done[0] = True
            root.after(0, on_solver_finished)

        except Exception as e:
            # This runs if CBC was killed mid-solve (e.g. user closed the window)
            # or if any other unexpected error occurred in the solver
            #
            # WHY CHECK winfo_exists()?
            # If the window is already destroyed (user closed it), trying to
            # update the label would cause a second crash. So we check first.
            solver_done[0] = True  # Stop the animation either way
            if root.winfo_exists():
                def show_error():
                    set_solving_state(False)
                    if animate_job[0] is not None:
                        root.after_cancel(animate_job[0])
                    error_label.config(text=f"Solver stopped: {str(e)}")
                root.after(0, show_error)
    # -------------------------------------------------------
    # PART 4: What happens when the solver finishes
    # -------------------------------------------------------
    def on_solver_finished():
        # Cancel any pending animation call
        if animate_job[0] is not None:
            root.after_cancel(animate_job[0])

        # RE-ENABLE everything now that solving is done
        set_solving_state(False)

        # Unpack the results
        assignments, summary = result_holder[0]

        # Everything below is exactly your original create_rota code
        # — just moved here so it runs after the solver is done

        print("Final Rota:")
        for shift_name, worker in assignments.items():
            print(f"{shift_name}: {worker if worker else 'Unassigned'}")

        if summary["status"] == "Infeasible":
            error_label.config(text="ERROR: Impossible to create rota with current rules! Check shift ranges, max weekends, max 24hr shifts.")
            return

        if summary["status"] == "Nothing to assign":
            error_label.config(text="No empty shifts or no workers — nothing to assign.")
            return

        popup = Toplevel(root)
        popup.title("Rota Results")

        text_widget = Text(popup, wrap="none", width=60, height=30, font=("Courier", 10))
        text_widget.pack(side="left", fill="both", expand=True)

        scrollbar = Scrollbar(popup, orient="vertical", command=text_widget.yview)
        scrollbar.pack(side="right", fill="y")
        text_widget.config(yscrollcommand=scrollbar.set)

        def on_popup_mousewheel(event):
            text_widget.yview_scroll(int(-1 * (event.delta / 120)), "units")

        popup.bind("<MouseWheel>", on_popup_mousewheel)
        text_widget.bind("<MouseWheel>", on_popup_mousewheel)

        text_widget.insert("end", "Final Rota (Multi-Unit)\n", "title")
        text_widget.insert("end", "=" * 70 + "\n\n", "separator")

        """
        Below a dictionary is built for assignments_by_unit = {
        "ICU": {
            5:  {"Day": "Dr. Smith", "Night": "Dr. Jones"},
            12: {"Day": "Unassigned", "Night": "Dr. Smith"},
            ...},
        "Internal Medicine": {
            2: {"Day": "Bobby", "Night": "Dr. Brown"} 
            ...},
        }
        """
        assignments_by_unit = {}
        for shift_name, worker in assignments.items():
            if worker is None:
                worker = "Unassigned" # Later "No shift" is overwritten -> to "Unassigned", but only if the solver did not find a worker. If the shift did not even exist, it remains as "No shift", which is written as default for every day (1, 2...) below.
            unit = extract_unit_from_shift_name(shift_name)
            if unit not in assignments_by_unit:
                assignments_by_unit[unit] = {}
            day = extract_day_from_shift_name(shift_name)
            shift_type = shift_name.split()[0]
            if day not in assignments_by_unit[unit]:
                assignments_by_unit[unit][day] = {"Day": "No shift", "Night": "No shift"} # First builds a "No shift", so it can be filled by a worker later.
            assignments_by_unit[unit][day][shift_type] = worker # Here "Unassigned" could overwrite "No shift", or a specific worker like "Dr. Brown" overwrites "No shift"

        for unit in sorted(assignments_by_unit.keys()):
            text_widget.insert("end", f"=== {unit} ===\n", "unit_header")
            text_widget.insert("end", "-" * 70 + "\n", "separator")
            text_widget.insert("end", f"{'Day':<7}{'Day Shift':<30}{'Night Shift':<30}\n", "header")
            text_widget.insert("end", "-" * 70 + "\n", "separator")
            for day in sorted(assignments_by_unit[unit].keys()):
                day_worker = assignments_by_unit[unit][day].get("Day", "No shift")
                night_worker = assignments_by_unit[unit][day].get("Night", "No shift")
                text_widget.insert("end", f"{day:<7}{day_worker:<30}{night_worker:<30}\n")
            text_widget.insert("end", "\n")

        text_widget.insert("end", "=" * 70 + "\n", "separator")
        text_widget.insert("end", "Summary:\n", "title")
        text_widget.insert("end", "-" * 70 + "\n", "separator")
        text_widget.insert("end", f"Number of preferred shifts assigned: {summary['preferences_count']}\n")
        text_widget.insert("end", f"Number of 24-hour shifts: {summary['twenty_four_count']}\n")
        text_widget.insert("end", f"Number of bad spacing pairs (<{spacing_days_threshold} days apart): {summary['bad_spacing_count']}\n")
        text_widget.insert("end", f"Number of shifts in non-preferred units (workers with preference only): {summary['non_preferred_unit_count']}\n")

        text_widget.tag_config("title", font=("Courier", 12, "bold"))
        text_widget.tag_config("unit_header", font=("Courier", 11, "bold"), foreground="blue")
        text_widget.tag_config("header", font=("Courier", 10, "bold"))
        text_widget.tag_config("separator", foreground="gray")

        text_widget.config(state="disabled")
        error_label.config(text="Rota finished! See the popup window for results.")

    # -------------------------------------------------------
    # PART 5: Kick everything off
    # -------------------------------------------------------
    # Start the animation immediately
    animate()

    # Start the solver on a background thread
    # daemon=True means: if the user closes the window,
    # don't let this thread keep Python alive in the background
    solver_thread = threading.Thread(target=run_solver, daemon=True)
    solver_thread.start()

# ---------------------------------------------------------------
# Small code area mainly for buttons and error sign at the bottom
# ---------------------------------------------------------------

# Shared frame so both buttons sit side by side
action_frame = Frame(root)
action_frame.pack(pady=4)

add_worker_button = Button(action_frame, text="➕ Add Worker", command=add_worker_row, width=12)
add_worker_button.pack(side=LEFT, padx=4)

create_rota_button = Button(action_frame, text="Create Rota ⭕", command=create_rota, width=12)
create_rota_button.pack(side=LEFT, padx=4)

Button(utilities_lf, text="PuLP Settings",    width=16, command=open_pulp_settings).pack(pady=3)
Button(utilities_lf, text="Save Preferences", width=16, command=save_preferences).pack(pady=3)
Button(utilities_lf, text="Load",             width=16, command=load_preferences).pack(pady=3)
Button(utilities_lf, text="Load .xlsx",       width=16, command=load_xlsx_preferences).pack(pady=3)

# Error label (same).
error_label = Label(root, text="")  # For errors.
error_label.pack()

def on_closing():
    """
    This runs when the user clicks the X to close the window.
    
    WHY NEEDED?
    CBC runs as a child OS process. When Python exits normally,
    child processes are NOT automatically killed - they keep running.
    We need to manually find and kill them.
    
    psutil lets us find all child processes of our own Python process
    and kill them before we exit. Subprocess let's also kill the whole tree.
    """
    try:
        current_process = psutil.Process(os.getpid())
        children = current_process.children(recursive=True)

        for child in children:
            try:
                print(f"Killing child process: {child.name()} (PID {child.pid})")

                # /F = force kill, /T = kill the whole tree including conhost.exe
                subprocess.call(
                    ['taskkill', '/F', '/T', '/PID', str(child.pid)],
                    creationflags=subprocess.CREATE_NO_WINDOW
                )

            except psutil.NoSuchProcess:
                pass
            except Exception:
                # Fallback for macOS/Linux where taskkill doesn't exist
                try:
                    child.kill()
                except psutil.NoSuchProcess:
                    pass

    except Exception as e:
        print(f"Cleanup error: {e}")

    root.destroy() # Now close the window

# Tell Tkinter to run on_closing instead of just closing
# WM_DELETE_WINDOW is the event that fires when user clicks X
root.protocol("WM_DELETE_WINDOW", on_closing)

root.mainloop()  # Start the window – like "go!"