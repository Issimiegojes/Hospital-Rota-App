"""
SOLVER.PY - The Brain of the Rota App
=====================================

This file does ONE job: solve the rota scheduling problem.
It's like a puzzle solver that figures out who works when.

Think of it like this:
- main.py = The face of the app (buttons, windows, what you see)
- solver.py = The brain (the math and logic that figures out the schedule)
"""

import pulp  # This is the puzzle-solving library
import sys   # Helps us find files when running as .exe
import os    # Helps us work with file paths


def solve_rota(shifts_list, workers_list, settings):
    """
    CREATE_ROTA - The main function that solves the scheduling puzzle
    
    What it does:
    - Takes a list of shifts (like "Day 1", "Night 15")
    - Takes a list of workers (doctors/nurses with their preferences)
    - Figures out the best way to assign workers to shifts
    
    Parameters (inputs):
    - shifts_list: A list of dictionaries, each representing a shift
      Example: {"name": "Day 1", "type": "Day", "tags": ["Monday"], "assigned_worker": None}
    - workers_list: A list of dictionaries, each representing a worker
      Example: {"name": "John", "shifts_to_fill": [3, 5], "cannot_work": ["Day 1"], ...}
    - settings: A dictionary with scoring rules
      Example: {"points_filled": 100, "points_preferred": 5, ...}
    
    Returns:
    - assignments: A dictionary mapping shift names to worker names
      Example: {"Day 1": "John", "Night 1": "Sarah", ...}
    - summary: A dictionary with statistics
      Example: {"preferences_count": 15, "twenty_four_count": 3, ...}
    """
    
    # ============================================================================
    # STEP 1: Extract settings from the settings dictionary
    # ============================================================================
    # Think of settings like a recipe card - we're pulling out each ingredient
    
    points_filled = settings.get("points_filled", 100)
    points_preferred = settings.get("points_preferred", 5)
    points_spacing = settings.get("points_spacing", -1)
    spacing_days_threshold = settings.get("spacing_days_threshold", 5)
    points_24hr = settings.get("points_24hr", -10)
    enforce_no_adj_nights = settings.get("enforce_no_adj_nights", True)
    enforce_no_adj_days = settings.get("enforce_no_adj_days", True)
    
    # ============================================================================
    # STEP 2: Build the assignments dictionary from shifts_list
    # ============================================================================
    # This creates a mapping: shift name → worker (or None if unassigned)
    # Example: {"Day 1": "John", "Night 1": None, "Day 2": "Sarah"}
    
    assignments = {}
    for shift in shifts_list:
        # For each shift, store who's assigned (might be None)
        assignments[shift["name"]] = shift["assigned_worker"]
    
    # ============================================================================
    # STEP 3: Find empty shifts (the ones we need to fill)
    # ============================================================================
    # We only want to assign workers to shifts that are currently empty (None)
    
    empty_shifts = [name for name in assignments if assignments[name] is None]
    
    # Find empty weekend shifts specifically (for weekend limit rule)
    empty_weekend_shifts = []
    for shift in shifts_list:
        # If the shift is empty AND tagged as "Weekend", add it to the list
        if shift["name"] in empty_shifts and "Weekend" in shift["tags"]:
            empty_weekend_shifts.append(shift["name"])
    
    # ============================================================================
    # STEP 4: Get the list of worker names
    # ============================================================================
    # Extract just the names from the workers_list
    # This turns [{"name": "John", ...}, {"name": "Sarah", ...}]
    # Into ["John", "Sarah"]
    
    workers = [w["name"] for w in workers_list]
    
    # ============================================================================
    # STEP 5: Find "bad pairs" of shifts (shifts that shouldn't be consecutive)
    # ============================================================================
    # These are rules like "don't work Night then Day the next day"
    
    bad_night_to_day_pairs = []      # Night shift followed by Day shift next day
    bad_adjacent_nights_pairs = []   # Night shift followed by Night shift next day
    bad_adjacent_days_pairs = []     # Day shift followed by Day shift next day
    twenty_four_hour_shift_pairs = [] # Day shift AND Night shift on the SAME day
    
    # Sort the empty shifts by day number (so we can check consecutive days)
    # This turns ["Day 5", "Night 2", "Day 3"] into ["Night 2", "Day 3", "Day 5"]
    sorted_shift_names = sorted(empty_shifts, key=lambda s: int(s.split(" ")[1]))
    
    # Loop through all pairs of shifts
    for i in range(len(sorted_shift_names)):
        for j in range(i+1, len(sorted_shift_names)):
            current = sorted_shift_names[i]      # Example: "Night 5"
            next_shift = sorted_shift_names[j]   # Example: "Day 6"
            
            # Extract the day numbers from the shift names
            # "Night 5" → 5, "Day 6" → 6
            current_day = int(current.split(" ")[1])
            next_day = int(next_shift.split(" ")[1])
            
            # Check if these are consecutive days (next_day = current_day + 1)
            if next_day == current_day + 1:
                # Rule 1: Night → Day (not allowed)
                if "Night" in current and "Day" in next_shift:
                    bad_night_to_day_pairs.append((current, next_shift))
                
                # Rule 2: Night → Night (not allowed if setting is True)
                if "Night" in current and "Night" in next_shift:
                    bad_adjacent_nights_pairs.append((current, next_shift))
                
                # Rule 3: Day → Day (not allowed if setting is True)
                if "Day" in current and "Day" in next_shift:
                    bad_adjacent_days_pairs.append((current, next_shift))
            
            # Rule 4: Same day (24-hour shift - discouraged with negative points)
            if next_day == current_day:
                twenty_four_hour_shift_pairs.append((current, next_shift))
    
    # ============================================================================
    # STEP 6: Find shifts that are too close together (bad spacing)
    # ============================================================================
    # Example: Working on Day 3 and Day 7 is only 4 days apart - too close!
    
    # Get ALL shift names (not just empty ones) and sort by day
    all_shift_names = sorted([shift["name"] for shift in shifts_list], 
                            key=lambda s: int(s.split(" ")[1]))
    
    bad_spacing_pairs = []
    
    for i in range(len(all_shift_names)):
        for j in range(i+1, len(all_shift_names)):
            shift1 = all_shift_names[i]
            shift2 = all_shift_names[j]
            
            day1 = int(shift1.split(" ")[1])
            day2 = int(shift2.split(" ")[1])
            
            # If shifts are too close (less than threshold days apart)
            # AND both are empty (available to assign)
            if day2 - day1 < spacing_days_threshold:
                if shift1 in empty_shifts and shift2 in empty_shifts:
                    bad_spacing_pairs.append((shift1, shift2))
    
    # ============================================================================
    # STEP 7: Extract worker preferences and limits
    # ============================================================================
    # Create dictionaries (like quick lookup tables) for each worker's info
    
    # What shifts each worker PREFERS to work
    worker_prefers = {w["name"]: w["prefers"] for w in workers_list}
    
    # How many 24-hour shifts each worker can do (maximum)
    max_24hr = {w["name"]: w["max_24hr"] for w in workers_list}
    
    # How many weekend shifts each worker can do (maximum)
    max_weekends = {w["name"]: w["max_weekends"] for w in workers_list}
    
    # What shifts each worker CANNOT work (hard constraint)
    worker_cannot = {w["name"]: w["cannot_work"] for w in workers_list}
    
    # ============================================================================
    # STEP 8: Check if there's anything to solve
    # ============================================================================
    # If there are no empty shifts OR no workers, we can't assign anything!
    
    if not empty_shifts or not workers:
        print("No empty shifts or no workers – nothing to assign.")
        return assignments, {
            "preferences_count": 0,
            "twenty_four_count": 0,
            "bad_spacing_count": 0,
            "status": "Nothing to assign"
        }
    
    # ============================================================================
    # STEP 9: Create the optimization problem
    # ============================================================================
    # This is where we set up the "puzzle" for PuLP to solve
    
    # Create a problem instance - we want to MAXIMIZE our score
    prob = pulp.LpProblem("Rota_Assignment", pulp.LpMaximize)
    
    # ============================================================================
    # STEP 10: Create decision variables
    # ============================================================================
    # These are the "unknowns" that PuLP will figure out
    
    # assign_vars[worker][shift] = 1 if worker is assigned to shift, 0 otherwise
    # Example: assign_vars["John"]["Day 5"] could be 0 or 1
    assign_vars = pulp.LpVariable.dicts("Assign", 
                                       (workers, empty_shifts), 
                                       0, 1, 
                                       pulp.LpBinary)
    
    # twenty_four_vars[worker][pair] = 1 if worker does both shifts in the pair
    # This helps us track 24-hour shifts
    twenty_four_vars = pulp.LpVariable.dicts("24hr", 
                                            (workers, twenty_four_hour_shift_pairs), 
                                            0, 1, 
                                            pulp.LpBinary)
    
    # spacing_var[worker][pair] = 1 if worker does both shifts in a bad spacing pair
    # This helps us penalize shifts that are too close together
    spacing_var = pulp.LpVariable.dicts("SpacingBad", 
                                       (workers, bad_spacing_pairs), 
                                       0, 1, 
                                       pulp.LpBinary)
    
    # ============================================================================
    # STEP 11: Define the objective function (what we want to maximize)
    # ============================================================================
    # This is like a scoring system - we want the highest score possible!
    
    prob += (
        # Points for each shift we fill (positive - we WANT to fill shifts)
        points_filled * pulp.lpSum(assign_vars[w][shift] 
                                   for w in workers 
                                   for shift in empty_shifts)
        
        # Extra points for filling preferred shifts (positive - workers are happier)
        + points_preferred * pulp.lpSum(assign_vars[w][shift] 
                                       for w in workers 
                                       for shift in empty_shifts 
                                       if shift in worker_prefers[w])
        
        # Penalty for shifts that are too close together (negative - we DON'T want this)
        + points_spacing * pulp.lpSum(spacing_var[w][pair] 
                                     for w in workers 
                                     for pair in bad_spacing_pairs)
        
        # Penalty for 24-hour shifts (negative - we want to minimize these)
        + points_24hr * pulp.lpSum(twenty_four_vars[w][pair] 
                                  for w in workers 
                                  for pair in twenty_four_hour_shift_pairs)
    )
    
    # ============================================================================
    # STEP 12: Add constraints (rules that MUST be followed)
    # ============================================================================
    
    # CONSTRAINT 1: Each shift can have at most 1 worker
    # (We can't have two people doing the same shift!)
    for shift in empty_shifts:
        prob += pulp.lpSum(assign_vars[w][shift] for w in workers) <= 1
    
    # CONSTRAINT 2: Each worker must work within their shift range
    # Example: If a worker wants 3-5 shifts, they must work between 3 and 5
    for w in workers:
        # Find this worker's min and max shifts
        min_shifts, max_shifts = next(worker["shifts_to_fill"] 
                                     for worker in workers_list 
                                     if worker["name"] == w)
        
        # Worker can't work MORE than max_shifts
        prob += pulp.lpSum(assign_vars[w][shift] for shift in empty_shifts) <= max_shifts
        
        # Worker must work AT LEAST min_shifts
        prob += pulp.lpSum(assign_vars[w][shift] for shift in empty_shifts) >= min_shifts
    
    # CONSTRAINT 3: No worker can work Night then Day the next day
    # (This is a hard rule - it's physically exhausting!)
    for pair in bad_night_to_day_pairs:
        night, day = pair
        for w in workers:
            # The sum can't be 2 (both shifts), so it must be 0 or 1
            prob += assign_vars[w][night] + assign_vars[w][day] <= 1
    
    # CONSTRAINT 4: No adjacent night shifts (if enabled)
    # Night → Night next day is not allowed
    if enforce_no_adj_nights:
        for pair in bad_adjacent_nights_pairs:
            current, next_shift = pair
            for w in workers:
                prob += assign_vars[w][current] + assign_vars[w][next_shift] <= 1
    
    # CONSTRAINT 5: No adjacent day shifts (if enabled)
    # Day → Day next day is not allowed
    if enforce_no_adj_days:
        for pair in bad_adjacent_days_pairs:
            current, next_shift = pair
            for w in workers:
                prob += assign_vars[w][current] + assign_vars[w][next_shift] <= 1
    
    # CONSTRAINT 6: Link 24-hour variables
    # This ensures twenty_four_vars[w][pair] = 1 only when BOTH shifts are worked
    for pair in twenty_four_hour_shift_pairs:
        day, night = pair
        for w in workers:
            # If both shifts are 1, this makes twenty_four_vars >= 1, so it becomes 1
            prob += twenty_four_vars[w][pair] >= assign_vars[w][day] + assign_vars[w][night] - 1
            
            # If day shift is 0, this forces twenty_four_vars to be 0
            prob += twenty_four_vars[w][pair] <= assign_vars[w][day]
            
            # If night shift is 0, this forces twenty_four_vars to be 0
            prob += twenty_four_vars[w][pair] <= assign_vars[w][night]
    
    # CONSTRAINT 7: Limit on 24-hour shifts per worker
    for w in workers:
        prob += pulp.lpSum(twenty_four_vars[w][pair] 
                          for pair in twenty_four_hour_shift_pairs) <= max_24hr[w]
    
    # CONSTRAINT 8: Limit on weekend shifts per worker
    for w in workers:
        prob += pulp.lpSum(assign_vars[w][shift] 
                          for shift in empty_weekend_shifts) <= max_weekends[w]
    
    # CONSTRAINT 9: Link spacing variables
    # Similar to 24-hour variables, but for shifts that are too close together
    for pair in bad_spacing_pairs:
        s1, s2 = pair
        for w in workers:
            prob += spacing_var[w][pair] >= assign_vars[w][s1] + assign_vars[w][s2] - 1
            prob += spacing_var[w][pair] <= assign_vars[w][s1]
            prob += spacing_var[w][pair] <= assign_vars[w][s2]
    
    # CONSTRAINT 10: Workers cannot be assigned to shifts they marked as "cannot work"
    for w in workers:
        for forbidden_shift in worker_cannot[w]:
            if forbidden_shift in empty_shifts:
                # Force this to be 0 (not assigned)
                prob += assign_vars[w][forbidden_shift] <= 0
    
    # ============================================================================
    # STEP 13: Find the CBC solver path
    # ============================================================================
    # CBC is the actual solver program that does the optimization
    # We need to tell PuLP where to find it
    
    if hasattr(sys, '_MEIPASS'):
        # We're running as a .exe file (PyInstaller packages the app)
        cbc_path = os.path.join(sys._MEIPASS, 'cbc.exe')
    else:
        # We're running as a normal .py file
        cbc_path = 'cbc.exe'
    
    print(f"Using CBC path: {cbc_path}")
    
    # ============================================================================
    # STEP 14: Solve the problem!
    # ============================================================================
    print("Starting PuLP solve – time limit 60 seconds...")
    
    # Try to solve the problem with a 60-second time limit
    # NOTE: Choose one of these lines depending on your setup:
    
    # OPTION 1: If you have cbc.exe bundled with your app (for .exe distribution)
    # status = prob.solve(pulp.COIN_CMD(msg=1, timeLimit=60, path=cbc_path))
    
    # OPTION 2: If you're running the .py file directly (development mode)
    status = prob.solve(pulp.PULP_CBC_CMD(msg=1, timeLimit=60))
    
    # ============================================================================
    # STEP 15: Check if the solution is valid
    # ============================================================================
    
    if pulp.LpStatus[status] == "Infeasible":
        # The constraints are impossible to satisfy!
        # Example: Asking someone to work 10 shifts but only 5 shifts exist
        print("INFEASIBLE: Cannot create a valid rota with these constraints.")
        return assignments, {
            "preferences_count": 0,
            "twenty_four_count": 0,
            "bad_spacing_count": 0,
            "status": "Infeasible"
        }
    
    elif pulp.LpStatus[status] == "Optimal":
        print("Found best solution in time!")
    
    elif pulp.LpStatus[status] == "Not Solved":
        print("Timed out – showing best found so far.")
    
    else:
        print("Other issue:", pulp.LpStatus[status])
        return assignments, {
            "preferences_count": 0,
            "twenty_four_count": 0,
            "bad_spacing_count": 0,
            "status": pulp.LpStatus[status]
        }
    
    print(f"Solve finished! Status: {pulp.LpStatus[status]}")
    
    # ============================================================================
    # STEP 16: Extract the solution
    # ============================================================================
    # PuLP has found values for all our variables - now we read them
    
    for w in workers:
        for shift in empty_shifts:
            # If the variable = 1, assign this worker to this shift
            if assign_vars[w][shift].value() == 1:
                assignments[shift] = w
    
    print("Assignments done!")
    
    # Print the total score
    total_points = prob.objective.value()
    print("Total points for this rota:", total_points)
    
    # ============================================================================
    # STEP 17: Calculate summary statistics
    # ============================================================================
    # Count how many preferred shifts, 24-hour shifts, and bad spacing pairs
    
    preferences_count = 0
    twenty_four_count = 0
    bad_spacing_count = 0
    
    # Count preferred shifts
    for shift, worker in assignments.items():
        if worker is None:
            continue
        if shift in worker_prefers[worker]:
            preferences_count += 1
    
    # Group shifts by worker to count 24-hour and spacing issues
    worker_shifts = {}
    for shift, worker in assignments.items():
        if worker is not None:
            if worker not in worker_shifts:
                worker_shifts[worker] = []
            worker_shifts[worker].append(shift)
    
    # Count 24-hour shifts and bad spacing
    for worker, shifts in worker_shifts.items():
        # Sort shifts by day number
        sorted_shifts = sorted(shifts, key=lambda s: int(s.split(" ")[1]))
        
        for i in range(len(sorted_shifts) - 1):
            current = sorted_shifts[i]
            next_shift = sorted_shifts[i+1]
            
            current_day = int(current.split(" ")[1])
            next_day = int(next_shift.split(" ")[1])
            
            # Same day = 24-hour shift
            if current_day == next_day:
                twenty_four_count += 1
            
            # Too close = bad spacing
            if next_day - current_day < spacing_days_threshold:
                bad_spacing_count += 1
    
    # ============================================================================
    # STEP 18: Print final summary
    # ============================================================================
    print("Summary:")
    print("Number of preferred shifts assigned:", preferences_count)
    print("Number of 24-hour shifts:", twenty_four_count)
    print(f"Number of bad spacing pairs (<{spacing_days_threshold} days apart):", bad_spacing_count)
    
    # ============================================================================
    # STEP 19: Return results
    # ============================================================================
    summary = {
        "preferences_count": preferences_count,
        "twenty_four_count": twenty_four_count,
        "bad_spacing_count": bad_spacing_count,
        "status": pulp.LpStatus[status],
        "total_points": total_points
    }
    
    return assignments, summary