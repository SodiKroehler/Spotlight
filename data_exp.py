import pandas as pd


if __name__ == "__main__":
    file_path = 'combined.csv'
    combined_df = pd.read_csv(file_path)
    new_df = combined_df[~(combined_df['SheetName'] == 'All Near') & ~(combined_df['routing'] == 'exact')]
    new_df.to_csv('filtered.csv', index=False)