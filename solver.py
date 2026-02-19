"""
SOLVER.PY - The Brain of the Rota App (Multi-Unit Version)
===========================================================

This file solves the rota scheduling problem for multiple hospital units.
"""

import pulp
import sys
import os

def solve_rota(shifts_list, workers_list, units_list, settings):
    """
    Multi-unit rota solver.
    
    Handles shifts across multiple units with rules:
    - Only one shift per worker per day (except 24hr = Day+Night SAME unit)
    - No Night → Day next day (any units)
    - No Night → Night next day (any units)
    - No Day → Day next day (if enabled)
    - 24hr shifts only possible within same unit
    """
    
    # ============================================================================
    # STEP 1: Extract settings
    # ============================================================================
    points_filled = settings.get("points_filled", 100) # The .get method is used instead of settings["points_filled"] because it provides a fallback value: 100, in case the value doesn't exist yet
    points_preferred = settings.get("points_preferred", 5)
    points_preferred_unit = settings.get("points_preferred_unit", 5)
    points_spacing = settings.get("points_spacing", -1)
    spacing_days_threshold = settings.get("spacing_days_threshold", 5)
    points_24hr = settings.get("points_24hr", -10)
    enforce_no_adj_nights = settings.get("enforce_no_adj_nights", True)
    enforce_no_adj_days = settings.get("enforce_no_adj_days", True)
    
    # ============================================================================
    # STEP 2: Build assignments dictionary
    # ============================================================================
    assignments = {}
    for shift in shifts_list:
        assignments[shift["name"]] = shift["assigned_worker"]
    
    # ============================================================================
    # STEP 3: Find empty shifts
    # ============================================================================
    empty_shifts = [name for name in assignments if assignments[name] is None]
    
    # Find empty weekend shifts
    empty_weekend_shifts = []
    for shift in shifts_list:
        if shift["name"] in empty_shifts and "Weekend" in shift["tags"]:
            empty_weekend_shifts.append(shift["name"])
    
    # ============================================================================
    # STEP 4: Get worker names
    # ============================================================================
    workers = [w["name"] for w in workers_list]
    
    # ============================================================================
    # STEP 5: Helper function to parse shift names
    # ============================================================================
    def parse_shift_name(shift_name):
        """
        Parse shift name into components.
        
        Examples:
        - "Day 5 Cardiology" → ("Day", 5, "Cardiology")
        - "Night 12 Internal Medicine" → ("Night", 12, "Internal Medicine")
        """
        parts = shift_name.split()
        shift_type = parts[0]  # "Day" or "Night"
        day = int(parts[1])    # Day number
        unit = " ".join(parts[2:])  # Unit name (may contain spaces)
        return shift_type, day, unit
    
    # ============================================================================
    # STEP 6: Build bad pairs with multi-unit logic
    # ============================================================================
    
    bad_night_to_day_pairs = []       # Night (any unit) day X → Day (any unit) day X+1
    bad_adjacent_nights_pairs = []    # Night (any unit) day X → Night (any unit) day X+1
    bad_adjacent_days_pairs = []      # Day (any unit) day X → Day (any unit) day X+1
    twenty_four_hour_shift_pairs = [] # Day X unit A + Night X unit A (SAME unit only)
    bad_same_day_non24hr_pairs = []   # Any two shifts on same day that are NOT (Day+Night same unit)
    
    # Group shifts by day for easier processing
    shifts_by_day = {}
    for shift_name in empty_shifts:
        shift_type, day, unit = parse_shift_name(shift_name)
        if day not in shifts_by_day:
            shifts_by_day[day] = []
        shifts_by_day[day].append((shift_name, shift_type, unit))
    
    # ========================================================================
    # Process consecutive days for Night→Day, Night→Night, Day→Day
    # ========================================================================
    for day in sorted(shifts_by_day.keys()):
        if day + 1 in shifts_by_day:
            # Compare all shifts on day X with all shifts on day X+1
            for shift1_name, type1, unit1 in shifts_by_day[day]:
                for shift2_name, type2, unit2 in shifts_by_day[day + 1]:
                    # Rule: Night → Day (any units) is forbidden
                    if type1 == "Night" and type2 == "Day":
                        bad_night_to_day_pairs.append((shift1_name, shift2_name))
                    
                    # Rule: Night → Night (any units) is forbidden
                    if type1 == "Night" and type2 == "Night":
                        bad_adjacent_nights_pairs.append((shift1_name, shift2_name))
                    
                    # Rule: Day → Day (any units) is forbidden (if enabled)
                    if type1 == "Day" and type2 == "Day":
                        bad_adjacent_days_pairs.append((shift1_name, shift2_name))
    
    # ========================================================================
    # Process same-day shifts
    # ========================================================================
    for day, shifts_on_day in shifts_by_day.items():
        # Compare all pairs of shifts on the same day
        for i, (shift1_name, type1, unit1) in enumerate(shifts_on_day):
            for shift2_name, type2, unit2 in shifts_on_day[i+1:]:
                # Check if this is a valid 24hr shift (Day + Night, SAME unit)
                is_same_unit = (unit1 == unit2)
                is_day_and_night = {type1, type2} == {"Day", "Night"}
                
                if is_same_unit and is_day_and_night:
                    # This is an ALLOWED 24hr shift
                    twenty_four_hour_shift_pairs.append((shift1_name, shift2_name))
                else:
                    # This is FORBIDDEN (different units or both same type)
                    # Examples: 
                    # - "Day 5 Cardiology" + "Day 5 Internal Medicine" (different units)
                    # - "Day 5 Cardiology" + "Night 5 Internal Medicine" (different units)
                    # - "Day 5 Cardiology" + "Day 5 Cardiology" (same type, impossible but handle it)
                    bad_same_day_non24hr_pairs.append((shift1_name, shift2_name))
    
    # ========================================================================
    # Build spacing pairs (shifts too close together)
    # ========================================================================
    all_shift_names = sorted([shift["name"] for shift in shifts_list], 
                            key=lambda s: parse_shift_name(s)[1])  # Sort by day number
    
    bad_spacing_pairs = []
    for i in range(len(all_shift_names)):
        for j in range(i+1, len(all_shift_names)):
            shift1 = all_shift_names[i]
            shift2 = all_shift_names[j]
            
            _, day1, _ = parse_shift_name(shift1)
            _, day2, _ = parse_shift_name(shift2)
            
            # If shifts are too close together AND both are empty
            if day2 - day1 < spacing_days_threshold:
                if shift1 in empty_shifts and shift2 in empty_shifts:
                    bad_spacing_pairs.append((shift1, shift2))
    
    # ============================================================================
    # STEP 7: Extract worker preferences and limits
    # ============================================================================
    # Convert from {row_num: ["Cardiology"]} to {"Dr. Smith": ["Cardiology"]}
    worker_preferred_units = {}
    for w in workers_list:
        worker_preferred_units[w["name"]] = w.get("prefer_units", [])
    
    max_24hr = {w["name"]: w["max_24hr"] for w in workers_list}
    max_weekends = {w["name"]: w["max_weekends"] for w in workers_list}

    worker_prefers = {}
    for w in workers_list:
        name = w["name"]
        old_prefers = w["prefers"]   # e.g. ["Day 5", "Night 12"]
        new_prefers = []       
        for unit in units_list:
            for old_shift in old_prefers:
                # old_shift is like "Day 5" or "Night 3"
                new_shift = f"{old_shift} {unit}"
                new_prefers.append(new_shift)
        worker_prefers[name] = new_prefers # new shift is "Day 5 Cardiology", "Day 5 Internal Medicine"

    # Now transform the cannot lists
    worker_cannot = {}
    for w in workers_list:
        name = w["name"]
        old_forbidden = w["cannot_work"]   # e.g. ["Day 5", "Night 12"]
        new_forbidden = []       
        for unit in units_list:
            for old_shift in old_forbidden:
                # old_shift is like "Day 5" or "Night 3"
                new_shift = f"{old_shift} {unit}"
                new_forbidden.append(new_shift)
        worker_cannot[name] = new_forbidden

    # ============================================================================
    # STEP 8: Check if there's anything to solve
    # ============================================================================
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
    prob = pulp.LpProblem("Rota_Assignment_MultiUnit", pulp.LpMaximize)
    
    # ============================================================================
    # STEP 10: Create decision variables
    # ============================================================================
    assign_vars = pulp.LpVariable.dicts("Assign", 
                                       (workers, empty_shifts), 
                                       0, 1, 
                                       pulp.LpBinary)
    
    twenty_four_vars = pulp.LpVariable.dicts("24hr", 
                                            (workers, twenty_four_hour_shift_pairs), 
                                            0, 1, 
                                            pulp.LpBinary)
    
    spacing_var = pulp.LpVariable.dicts("SpacingBad", 
                                       (workers, bad_spacing_pairs), 
                                       0, 1, 
                                       pulp.LpBinary)
    
    # ============================================================================
    # STEP 11: Define the objective function
    # ============================================================================
    prob += (
        points_filled * pulp.lpSum(assign_vars[w][shift] 
                                   for w in workers 
                                   for shift in empty_shifts)
        + points_preferred * pulp.lpSum(assign_vars[w][shift] 
                                       for w in workers 
                                       for shift in empty_shifts 
                                       if shift in worker_prefers[w])
        + points_preferred_unit * pulp.lpSum(assign_vars[w][shift]
                                        for w in workers
                                        for shift in empty_shifts
                                        if parse_shift_name(shift)[2] in worker_preferred_units[w])
        + points_spacing * pulp.lpSum(spacing_var[w][pair] 
                                     for w in workers 
                                     for pair in bad_spacing_pairs)
        + points_24hr * pulp.lpSum(twenty_four_vars[w][pair] 
                                  for w in workers 
                                  for pair in twenty_four_hour_shift_pairs)
    )
    
    # ============================================================================
    # STEP 12: Add constraints
    # ============================================================================
    
    # CONSTRAINT 1: Each shift has at most 1 worker
    for shift in empty_shifts:
        prob += pulp.lpSum(assign_vars[w][shift] for w in workers) <= 1
    
    # CONSTRAINT 2: Worker must work within their shift range
    for w in workers:
        min_shifts, max_shifts = next(worker["shifts_to_fill"] 
                                     for worker in workers_list 
                                     if worker["name"] == w)
        prob += pulp.lpSum(assign_vars[w][shift] for shift in empty_shifts) <= max_shifts
        prob += pulp.lpSum(assign_vars[w][shift] for shift in empty_shifts) >= min_shifts
    
    # CONSTRAINT 3: No Night → Day next day (any units)
    for pair in bad_night_to_day_pairs:
        night, day = pair
        for w in workers:
            prob += assign_vars[w][night] + assign_vars[w][day] <= 1
    
    # CONSTRAINT 4: No adjacent nights (if enabled)
    if enforce_no_adj_nights:
        for pair in bad_adjacent_nights_pairs:
            current, next_shift = pair
            for w in workers:
                prob += assign_vars[w][current] + assign_vars[w][next_shift] <= 1
    
    # CONSTRAINT 5: No adjacent days (if enabled)
    if enforce_no_adj_days:
        for pair in bad_adjacent_days_pairs:
            current, next_shift = pair
            for w in workers:
                prob += assign_vars[w][current] + assign_vars[w][next_shift] <= 1
    
    # CONSTRAINT 6: NEW - No two shifts on same day (except allowed 24hr)
    # This prevents: "Day 5 Cardiology" + "Night 5 Internal Medicine"
    # But allows: "Day 5 Cardiology" + "Night 5 Cardiology" (handled separately)
    for pair in bad_same_day_non24hr_pairs:
        shift1, shift2 = pair
        for w in workers:
            prob += assign_vars[w][shift1] + assign_vars[w][shift2] <= 1
    
    # CONSTRAINT 7: Link 24-hour variables (only for same-unit Day+Night pairs)
    for pair in twenty_four_hour_shift_pairs:
        s1, s2 = pair
        
        # Figure out which is Day and which is Night
        type1, _, _ = parse_shift_name(s1)
        type2, _, _ = parse_shift_name(s2)
        
        if type1 == "Day":
            day_shift, night_shift = s1, s2
        else:
            day_shift, night_shift = s2, s1
        
        for w in workers:
            # If both are assigned, 24hr var becomes 1
            prob += twenty_four_vars[w][pair] >= assign_vars[w][day_shift] + assign_vars[w][night_shift] - 1
            # If day is 0, 24hr var must be 0
            prob += twenty_four_vars[w][pair] <= assign_vars[w][day_shift]
            # If night is 0, 24hr var must be 0
            prob += twenty_four_vars[w][pair] <= assign_vars[w][night_shift]
    
    # CONSTRAINT 8: Limit on 24-hour shifts per worker
    for w in workers:
        prob += pulp.lpSum(twenty_four_vars[w][pair] 
                          for pair in twenty_four_hour_shift_pairs) <= max_24hr[w]
    
    # CONSTRAINT 9: Limit on weekend shifts per worker
    for w in workers:
        prob += pulp.lpSum(assign_vars[w][shift] 
                          for shift in empty_weekend_shifts) <= max_weekends[w]
    
    # CONSTRAINT 10: Link spacing variables
    for pair in bad_spacing_pairs:
        s1, s2 = pair
        for w in workers:
            prob += spacing_var[w][pair] >= assign_vars[w][s1] + assign_vars[w][s2] - 1
            prob += spacing_var[w][pair] <= assign_vars[w][s1]
            prob += spacing_var[w][pair] <= assign_vars[w][s2]
    
    # CONSTRAINT 11: Workers cannot be assigned to forbidden shifts
    for w in workers:
        for forbidden_shift in worker_cannot[w]:
            if forbidden_shift in empty_shifts:
                prob += assign_vars[w][forbidden_shift] <= 0
    
    # ============================================================================
    # STEP 13: Solve the problem
    # ============================================================================
    print("Starting PuLP solve – time limit 6000 seconds...")
    timeLimit_setting = 6000

    # Falls back to CBC if not .exe file:
    if hasattr(sys, '_MEIPASS'):
        cbc_path = os.path.join(sys._MEIPASS, 'cbc.exe')
        status = prob.solve(pulp.COIN_CMD(msg=1, timeLimit=timeLimit_setting, path=cbc_path))
    else:
        status = prob.solve(pulp.PULP_CBC_CMD(msg=1, timeLimit=timeLimit_setting))
       # try: this part uses highspy/highsolver, but crashes on "Presolving model"
       #     # Use highspy directly (Python API, no executable needed)
       #     solver = pulp.getSolver('HiGHS')
       #     status = prob.solve(solver)
       #     print("Used HiGHS solver via highspy")
       # except Exception as e:
       #     print(f"HiGHS failed — exact reason: {e}")
       #     status = prob.solve(pulp.PULP_CBC_CMD(msg=1, timeLimit=timeLimit_setting))

# ============================================================================
    # STEP 14: Check solution status
    # ============================================================================
    if pulp.LpStatus[status] == "Infeasible":
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
    # STEP 15: Extract the solution
    # ============================================================================
    for w in workers:
        for shift in empty_shifts:
            if assign_vars[w][shift].value() == 1:
                assignments[shift] = w
    
    print("Assignments done!")
    
    # Print total score
    total_points = prob.objective.value()
    print("Total points for this rota:", total_points)
    
    # ============================================================================
    # STEP 16: Calculate summary statistics
    # ============================================================================
    preferences_count = 0
    twenty_four_count = 0
    bad_spacing_count = 0
    
    # Count preferred shifts
    for shift, worker in assignments.items():
        if worker is None:
            continue
        if shift in worker_prefers[worker]:
            preferences_count += 1
    
    # Group shifts by worker
    worker_shifts = {}
    for shift, worker in assignments.items():
        if worker is not None:
            if worker not in worker_shifts:
                worker_shifts[worker] = []
            worker_shifts[worker].append(shift)
    
    # Count 24-hour shifts and bad spacing
    for worker, shifts in worker_shifts.items():
        # Sort shifts by day number
        sorted_shifts = sorted(shifts, key=lambda s: parse_shift_name(s)[1])
        
        for i in range(len(sorted_shifts) - 1):
            current = sorted_shifts[i]
            next_shift = sorted_shifts[i+1]
            
            _, current_day, current_unit = parse_shift_name(current)
            _, next_day, next_unit = parse_shift_name(next_shift)
            
            # Same day AND same unit = 24-hour shift
            if current_day == next_day and current_unit == next_unit:
                twenty_four_count += 1
            
            # Too close = bad spacing
            if next_day - current_day < spacing_days_threshold:
                bad_spacing_count += 1
    
    # ============================================================================
    # STEP 17: Print summary
    # ============================================================================
    print("Summary:")
    print("Number of preferred shifts assigned:", preferences_count)
    print("Number of 24-hour shifts:", twenty_four_count)
    print(f"Number of bad spacing pairs (<{spacing_days_threshold} days apart):", bad_spacing_count)
    
    # ============================================================================
    # STEP 18: Return results
    # ============================================================================
    summary = {
        "preferences_count": preferences_count,
        "twenty_four_count": twenty_four_count,
        "bad_spacing_count": bad_spacing_count,
        "status": pulp.LpStatus[status],
        "total_points": total_points
    }
    
    return assignments, summary
