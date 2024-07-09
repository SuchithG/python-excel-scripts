import pandas as pd

# Load the notification report
notification_report_path = 'path/to/notification_report_2024.xlsx'
notification_df = pd.read_excel(notification_report_path)

# Load the June 28 data
june_28_path = 'path/to/28 June.xlsx'
june_28_df = pd.read_excel(june_28_path)

# Sum the counts of each notification ID
exceptions_count_eod = june_28_df.groupby('NOTFCN_ID')['NOTFCN_CNT'].sum().reset_index()
exceptions_count_eod.columns = ['Notification', 'Exceptions count EOD']

# Merge the sum of counts with the notification report
updated_notification_df = pd.merge(notification_df, exceptions_count_eod, on='Notification', how='left')

# Save the modified notification report
updated_notification_df.to_excel('path/to/modified_notification_report_2024.xlsx', index=False)

print("New column 'Exceptions count EOD' added and file saved as 'modified_notification_report_2024.xlsx'")
