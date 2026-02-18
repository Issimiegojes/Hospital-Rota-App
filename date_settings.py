from tkinter import *
import calendar

# Function for save year (like button press).
def save_year_confirm(save_year_inputs):

    year_entry = save_year_inputs["give_year_entry"]
    error_label = save_year_inputs["give_error_label"]
    current_year_label = save_year_inputs["give_current_year_label"]
    year = save_year_inputs["give_year"]

    year_input = year_entry.get().strip()  # Get from type box and strip whitespace.
    try:
        year = int(year_input)
        if year < 1900 or year > 2100:
            error_label.config(text="Bad year – 1900-2100.")  # Show error.
        else:
            current_year_label.config(text="Current Year: " + str(year))  # Update show.
            error_label.config(text="")  # Clear error.
            return year
    except ValueError:
        error_label.config(text="Not a number!")

def save_month_confirm(save_month_inputs):  # Function for the button.
    
    month_entry = save_month_inputs["give_month_entry"]
    error_label = save_month_inputs["give_error_label"]
    current_month_label = save_month_inputs["give_current_month_label"]
    month = save_month_inputs["give_month"]
    year = save_month_inputs["give_year"]
    
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

        return month, num_days, starting_weekday, days_list
    except ValueError:
        error_label.config(text="That's not a number! Please try again.")  # Error.

def save_holidays_confirm(save_holidays_inputs):  # Function for the button.

    month = save_holidays_inputs["give_month"]
    year = save_holidays_inputs["give_year"]
    holiday_entry = save_holidays_inputs["give_holiday_entry"]
    holidays_label = save_holidays_inputs["give_holidays_label"]
    error_label = save_holidays_inputs["give_error_label"]

    month_details = calendar.monthrange(year, month)
    num_days = month_details[1]  # Days in month.

    holiday_input = holiday_entry.get()  # Get from type box.
    holiday_days = []  # Start empty.
    if holiday_input.strip() == "":  # If blank.
        holiday_days = []
        holidays_label.config(text="Holiday List: None")
        error_label.config(text="")  # Clear error.
        return holiday_days  # Done.
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
        return holiday_days
    except ValueError:
        error_label.config(text="Error: Some entries weren't numbers. Please try again.")  # Error.
