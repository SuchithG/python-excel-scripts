from datetime import datetime
import pandas as pd

def is_leap_year(year):
    return year % 4 == 0 and (year % 100 != 0 or year % 400 == 0)

def days_in_month(year, month):
    if month == 2:
        return 29 if is_leap_year(year) else 28
    elif month in [4, 6, 9, 11]:
        return 30
    else:
        return 31

def determine_age_category(creation_date, last_day_previous_month):
    year, month = last_day_previous_month.year, last_day_previous_month.month
    days_previous_month = days_in_month(year, month)

    age_days = (last_day_previous_month - creation_date).days

    if age_days <= 1: # 0-1 New
        return '0-1 New'
    elif age_days <= 7: # 2-7 Days
        return '02-07 days'
    elif age_days <= 15: # 8-15 Days
        return '08-15 days'
    elif age_days <= 30: # 16-30 Days
        return '16-30 days'
    elif age_days <= 180: # 31-180 Days
        return '31-180 days'
    else: # Beyond 180 days
        return '>180 days'

# Example usage
current_date = datetime.now().replace(day=1, hour=0, minute=0, second=0, microsecond=0)
last_day_previous_month = current_date - pd.Timedelta(days=1)

# Test with a date
test_date = datetime(2024, 1, 29)  # Example date
age_category = determine_age_category(test_date, last_day_previous_month)
print(age_category)
