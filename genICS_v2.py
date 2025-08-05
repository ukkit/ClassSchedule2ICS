import pandas as pd
from icalendar import Calendar, Event, Alarm
from datetime import datetime, timedelta
import pytz
import re


def create_ics_from_excel(excel_path, output_ics_path):
    # Read the Excel file
    df = pd.read_excel(excel_path, header=None)

    # Extract location from row 2
    location = df.iloc[1, 0] if pd.notna(df.iloc[1, 0]) else ""

    # Get the current week's Monday
    today = datetime.now(pytz.timezone('Asia/Kolkata'))
    monday = today - timedelta(days=today.weekday())

    # Create a calendar
    cal = Calendar()
    cal.add('prodid', '-//Class Timetable//mxm.dk//')
    cal.add('version', '2.0')

    # Days of the week
    days = ['Monday', 'Tuesday', 'Wednesday', 'Thursday', 'Friday']

    # Process each day's column
    for day_idx, day in enumerate(days):
        col = day_idx + 1  # Columns B to F (0-indexed would be 1 to 5)

        # Process each time slot
        row = 3  # Start from first period row
        while row < 25:  # Rows 4 to 24 (0-indexed would be 3 to 24)
            # Check if this is the start of a period (subject cell)
            subject_cell = df.iloc[row, col]
            if pd.isna(subject_cell) or str(subject_cell).strip() == "":
                row += 1
                continue

            # Get all 4 cells of the period
            subject = df.iloc[row, col]
            teacher = df.iloc[row+1, col] if row + \
                1 < len(df) and pd.notna(df.iloc[row+1, col]) else ""
            class_info = df.iloc[row+2, col] if row + \
                2 < len(df) and pd.notna(df.iloc[row+2, col]) else ""

            # Skip if essential info is missing
            if pd.isna(subject) or str(subject).strip() == "":
                row += 1
                continue

            # Get time slots for all 4 rows
            time_slots = []
            for i in range(4):
                if row + i >= len(df):
                    break
                time_slot = df.iloc[row + i, 0]
                if pd.isna(time_slot) or str(time_slot).strip() == "":
                    continue
                time_slots.append(str(time_slot).strip())

            if not time_slots:
                row += 1
                continue

            # Parse start time from first row's time slot
            first_time_match = re.match(
                r'(\d{1,2}:\d{2})', time_slots[0].split('-')[0])
            if not first_time_match:
                row += 1
                continue
            start_time_str = first_time_match.group(1)

            # Parse end time from last row's time slot
            last_time_slot = time_slots[-1] if time_slots else time_slots[0]
            last_time_match = re.match(
                r'.*?(\d{1,2}:\d{2})$', last_time_slot.split('-')[-1])
            if not last_time_match:
                row += 1
                continue
            end_time_str = last_time_match.group(1)

            # Parse times
            try:
                start_time = datetime.strptime(start_time_str, '%H:%M').time()
                end_time = datetime.strptime(end_time_str, '%H:%M').time()
            except ValueError:
                row += 4
                continue

            # Calculate event date (current week's day)
            event_date = monday + timedelta(days=day_idx)
            start_datetime = datetime.combine(event_date.date(), start_time)
            end_datetime = datetime.combine(event_date.date(), end_time)

            # Convert to timezone aware
            ist = pytz.timezone('Asia/Kolkata')
            start_datetime = ist.localize(start_datetime)
            end_datetime = ist.localize(end_datetime)

            # Create event title
            teacher_str = f" (by {teacher})" if teacher else ""
            event_title = f"{subject}{teacher_str}"

            # Create event
            event = Event()
            event.add('summary', event_title)
            event.add('location', location)
            event.add('dtstart', start_datetime)
            event.add('dtend', end_datetime)
            event.add('dtstamp', datetime.now(ist))

            # Add first alarm (30 minutes before)
            alarm30 = Alarm()
            alarm30.add('action', 'DISPLAY')
            alarm30.add('description',
                        f'Reminder: {event_title} at {location} in 30 minutes')
            alarm30.add('trigger', timedelta(minutes=-30))
            event.add_component(alarm30)

            # Add second alarm (15 minutes before)
            alarm15 = Alarm()
            alarm15.add('action', 'DISPLAY')
            alarm15.add('description',
                        f'Reminder: {event_title} at {location} in 15 minutes')
            alarm15.add('trigger', timedelta(minutes=-15))
            event.add_component(alarm15)

            cal.add_component(event)

            # Skip the next 3 rows as they're part of this event
            row += 4

    # Write to ICS file
    with open(output_ics_path, 'wb') as f:
        f.write(cal.to_ical())
    print(f"ICS file created at {output_ics_path}")


# Example usage
create_ics_from_excel('timetable.xlsx', 'class_schedule.ics')
