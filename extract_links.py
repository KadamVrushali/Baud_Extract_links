import pandas as pd
def extract_links_from_excel(file_path):
    """
    Reads all sheets of an Excel file and extracts links from the 'Link' column.

    Parameters:
        file_path (str): Path to the Excel file.

    Returns:
        tuple: (list_of_links, merged_dataframe)
    """

    # Read all sheets into a dict of DataFrames
    sheets = pd.read_excel(file_path, sheet_name=None, engine="openpyxl")

    all_dfs = []
    for sheet_name, df in sheets.items():
        # Normalize column names to lowercase for matching
        df.columns = [str(c).strip().lower() for c in df.columns]

        if "link" in df.columns:
            print(f"✅ Found 'Link' column in sheet: {sheet_name}")
            all_dfs.append(df)
        else:
            print(f"⚠ No 'Link' column in sheet: {sheet_name}")

    if not all_dfs:
        raise ValueError("No sheet contains a 'Link' column.")

    # Merge all sheets with Link column into one DataFrame
    merged_df = pd.concat(all_dfs, ignore_index=True)

    # Extract links (dropping NaNs)
    links = merged_df["link"].dropna().tolist()

    return links, merged_df
