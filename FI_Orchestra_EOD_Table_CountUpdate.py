import pandas as pd

# Load the notification report
notification_report_path = 'path/to/notification_report_2024.xlsx'
notification_df = pd.read_excel(notification_report_path)

# Load the June 28 data
june_28_path = 'path/to/28 June.xlsx'
june_28_df = pd.read_excel(june_28_path)

# Sum the notification counts
exceptions_count_eod = june_28_df['NOTFCN_CNT'].sum()

# Add the new column to the notification report
notification_df['Exceptions count EOD'] = exceptions_count_eod

# Save the modified notification report
notification_df.to_excel('path/to/modified_notification_report_2024.xlsx', index=False)

print("New column 'Exceptions count EOD' added and file saved as 'modified_notification_report_2024.xlsx'")
