import pandas as pd
import time
print ("H")

def csv_to_pkl(csv_file_path, pkl_file_path, chunksize=100000):
    """
    Converts a CSV file to a .pkl file efficiently using chunks.

    Parameters:
    csv_file_path (str): Path to the input CSV file.
    pkl_file_path (str): Path where the output .pkl file will be saved.
    chunksize (int): Number of rows to read at a time (default: 100000).

    Returns:
    None
    """
    # Initialize an empty DataFrame
    df = pd.DataFrame()

    # Read CSV in chunks to avoid memory issues
    for chunk in pd.read_csv(csv_file_path, chunksize=chunksize):
        df = pd.concat([df, chunk], ignore_index=True)

    # Save the DataFrame to a .pkl file
    df.to_pickle(pkl_file_path)
    print(f"CSV file successfully converted to {pkl_file_path}")


def pkl_to_excel(pkl_file_path, excel_file_path, chunksize=100000):
    """
    Unpickles data from a .pkl file and stores it in an Excel file efficiently,
    and displays the total conversion time.

    Parameters:
    pkl_file_path (str): Path to the input .pkl file.
    excel_file_path (str): Path where the output Excel file will be saved.
    chunksize (int): Number of rows to write at a time (default: 100000).

    Returns:
    None
    """
    # Start the timer
    start_time = time.time()

    # Load the entire DataFrame from the .pkl file
    df = pd.read_pickle(pkl_file_path)

    # Create an Excel writer object and specify the engine
    with pd.ExcelWriter(excel_file_path, engine='xlsxwriter') as writer:
        # Access the XlsxWriter workbook and worksheet objects
        workbook = writer.book
        worksheet = workbook.add_worksheet('Sheet1')

        # Add a format that turns off URL recognition
        url_format = workbook.add_format({'text_wrap': True, 'num_format': '@'})

        # Write the DataFrame to Excel in chunks
        for start_row in range(0, len(df), chunksize):
            # Define the end row for the current chunk
            end_row = min(start_row + chunksize, len(df))
            # Get the current chunk of data
            chunk = df[start_row:end_row]

            # Write the chunk to the Excel file without URL formatting
            chunk.to_excel(writer, index=False, header=(start_row == 0), startrow=start_row, sheet_name='Sheet1')

            # Apply the format to turn off URL recognition for long text fields
            worksheet.set_column(0, len(chunk.columns) - 1, None, url_format)

    # End the timer
    end_time = time.time()

    # Calculate the total conversion time
    total_time = end_time - start_time

    print(f".pkl file successfully converted to {excel_file_path}")
    print(f"Total conversion time: {total_time:.2f} seconds")

# Example usage
import pdb
pdb.set_trace()
# csv_file_path = 'C:/Users/akash/PycharmProjects/customers-2000000.csv'  # Replace with your CSV file path
#
# pkl_file_path = 'C:/Users/akash/PycharmProjects/large_data.pkl'  # Desired output .pkl file path
#
# csv_to_pkl(csv_file_path, pkl_file_path)

# Example usage
pkl_file_path = 'C:/Users/akash/PycharmProjects/large_data.pkl'  # Replace with your .pkl file path
excel_file_path = 'C:/Users/akash/PycharmProjects/large_data_output.xlsx'  # Desired output Excel file path

pkl_to_excel(pkl_file_path, excel_file_path)
