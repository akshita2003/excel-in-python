
import pandas as pd

# Load the data from the two sheets
df_users = pd.read_excel('C:/Users/AKSHITA SHARMA/Desktop/Book1.xlsx', sheet_name='Sheet1')
df_activity = pd.read_excel('C:/Users/AKSHITA SHARMA/Desktop/Book1.xlsx', sheet_name='Sheet2')

# Merge the two dataframes based on User ID
df_merged = pd.merge(df_users, df_activity, on='User ID')

# Calculate the average number of statements and reasons per team
team_stats = df_merged.groupby('Team Name').agg({'total_statements': 'mean', 'total_reasons': 'mean'})

# Calculate the average number of statements and reasons per user
user_stats = df_merged.groupby(['Team Name', 'Name']).agg({'total_statements': 'mean', 'total_reasons': 'mean'})

# Create a dataframe for the leaderboard
df_leaderboard = pd.DataFrame({
    'Team Rank': team_stats['total_statements'].rank(method='dense', ascending=False),
    'Thinking Teams Leaderboard': team_stats.index,
    'Average Statements': team_stats['total_statements'].round(2),
    'Average Reasons': team_stats['total_reasons'].round(2)
})

# Sort the leaderboard by rank
df_leaderboard = df_leaderboard.sort_values(by='Team Rank')

# Print the leaderboard
#print(df_leaderboard.to_string(index=False))
with pd.ExcelWriter('C:/Users/AKSHITA SHARMA/Desktop/Book1.xlsx', engine='openpyxl', mode='a') as writer:
    df_leaderboard.to_excel(writer, sheet_name='SHEET3',index=False)


# Calculate the total number of statements and reasons
df_merged['Total'] = df_merged['total_statements'] + df_merged['total_reasons']

# Sort the dataframe by Total, Name and UID
df_merged = df_merged.sort_values(['Total', 'Name', 'User ID'], ascending=[False, True, True])

# Add a Rank column based on the Total column
df_merged['Rank'] = df_merged['Total'].rank(method='min', ascending=False)

# Select the required columns for the leaderboard
leaderboard = df_merged[['Rank', 'Name', 'User ID', 'total_statements', 'total_reasons']]
leaderboard = leaderboard.rename(columns={'UID': 'User ID', 'total_statements': 'No. of Statements', 'total_reasons': 'No. of Reasons'})

# Store the resulting table into another sheet of the same file
with pd.ExcelWriter('C:/Users/AKSHITA SHARMA/Desktop/Book1.xlsx', engine='openpyxl', mode='a') as writer:
    leaderboard.to_excel(writer, sheet_name='Sheet4',index=False)

    



