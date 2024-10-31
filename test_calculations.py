import pandas as pd

def test_percentage_calculation(df):
    # Calculate total amount per PM Type
    pm_type_totals = df.groupby('PM Type')['Amount'].sum()
    
    # Group by PM Type and Main/CO/DCR
    grouped_df = df.groupby(['PM Type', 'Main/CO/DCR']).agg({
        'Amount': ['sum', 'count']
    }).reset_index()
    
    # Flatten column names
    grouped_df.columns = [
        col[0] if col[1] == '' else f"{col[0]}_{col[1]}" 
        for col in grouped_df.columns
    ]
    
    # Rename columns
    grouped_df = grouped_df.rename(columns={
        'Amount_sum': 'Total Amount',
        'Amount_count': 'Count'
    })
    
    # Calculate percentages
    grouped_df['Percentage of PM Type'] = grouped_df.apply(
        lambda row: (row['Total Amount'] / pm_type_totals[row['PM Type']]) * 100,
        axis=1
    )
    
    return grouped_df

# Test with your data
if __name__ == "__main__":
    df = pd.read_excel('summary_table_updated_analysis.xlsx', sheet_name='pc_overview AP')
    result = test_percentage_calculation(df)
    print(result)