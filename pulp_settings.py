from tkinter import *

def pulp_settings(root, settings_inputs, error_label, on_save):
    """
    Opens a popup window for editing PuLP solver settings.

    Arguments:
        root            - the main Tkinter window (parent)
        settings_inputs - dict of current setting values to pre-fill the form
        error_label     - the Label widget in the main window for status messages
        on_save         - callback function called with the new settings dict when Save is clicked
    """

    # Unpack the settings dictionary into local variables
    points_filled         = settings_inputs["points_filled"]
    points_preferred      = settings_inputs["points_preferred"]
    points_preferred_unit = settings_inputs["points_preferred_unit"]
    points_spacing        = settings_inputs["points_spacing"]
    spacing_days_threshold = settings_inputs["spacing_days_threshold"]
    points_24hr           = settings_inputs["points_24hr"]
    enforce_no_adj_days   = settings_inputs["enforce_no_adj_days"]
    enforce_no_adj_nights = settings_inputs["enforce_no_adj_nights"]
    include_weekday_days  = settings_inputs["include_weekday_days"]
    
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

    Label(popup, text="Points for preferred units:").grid(row=2, column=0, sticky="w")
    preferred_unit_entry = Entry(popup)
    preferred_unit_entry.insert(0, str(points_preferred_unit))
    preferred_unit_entry.grid(row=2, column=1)

    Label(popup, text="Points for bad spacing days:").grid(row=3, column=0, sticky="w")
    spacing_entry = Entry(popup)
    spacing_entry.insert(0, str(points_spacing))
    spacing_entry.grid(row=3, column=1)

    Label(popup, text="Days apart for spacing penalty:").grid(row=4, column=0, sticky="w")
    spacing_days_entry = Entry(popup)
    spacing_days_entry.insert(0, str(spacing_days_threshold))
    spacing_days_entry.grid(row=4, column=1)

    Label(popup, text="Points for 24-hour shifts:").grid(row=5, column=0, sticky="w")
    hr24_entry = Entry(popup)
    hr24_entry.insert(0, str(points_24hr))
    hr24_entry.grid(row=5, column=1)

    Label(popup, text="Enforce: No Day → Day").grid(row=6, column=0, sticky="w")
    adj_days_var = IntVar(value=1 if enforce_no_adj_days else 0)
    Checkbutton(popup, variable=adj_days_var).grid(row=6, column=1)

    Label(popup, text="Enforce: No Night → Night").grid(row=7, column=0, sticky="w")
    adj_nights_var = IntVar(value=1 if enforce_no_adj_nights else 0)
    Checkbutton(popup, variable=adj_nights_var).grid(row=7, column=1)

    Label(popup, text="Include Mon-Fri day shifts").grid(row=8, column=0, sticky="w")
    include_weekday_var = IntVar(value=1 if include_weekday_days else 0)
    Checkbutton(popup, variable=include_weekday_var).grid(row=8, column=1)

    def save_settings():
        try:
            new_settings = {
                    "points_filled":          int(filled_entry.get()),
                    "points_preferred":       int(preferred_entry.get()),
                    "points_preferred_unit":  int(preferred_unit_entry.get()),
                    "points_spacing":         int(spacing_entry.get()),
                    "spacing_days_threshold": int(spacing_days_entry.get()),
                    "points_24hr":            int(hr24_entry.get()),
                    "enforce_no_adj_days":    bool(adj_days_var.get()),
                    "enforce_no_adj_nights":  bool(adj_nights_var.get()),
                    "include_weekday_days":   bool(include_weekday_var.get()),
                }
            on_save(new_settings)  # send values back to hospital_rota_app.py
            error_label.config(text="PuLP settings updated successfully!")
            popup.destroy()
        except ValueError:
            error_label.config(text="Error: All values must be integers.")

    Button(popup, text="Save Settings", command=save_settings).grid(row=9, column=0, columnspan=2)
