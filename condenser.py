"""An Excel condenser tool which takes an Excel sheet with multiple nearly duplicate rows
and condenses these rows into a single row with extra columns. The condensed rows are then
saved to a new Excel file. 

@author: Ada Barach
@version: 02/03/2023
"""

import pandas as pd
import argparse
import numpy as np

def read_data_and_group(filepath):
    """Given a file path to an Excel document, reads in the data as a pandas.DataFrame
    object and returns the dataframe grouped by 'acct_id' column values. 

    Args:
        filepath (string): absolute file path to the Excel document

    Returns:
        pandas.DataFrame: a DataFrame containing the cells from the Excel worksheet, 
        grouped by account id. 
    """
    # read data
    df = pd.read_excel(filepath)
    
    # add handling charge if srvchg is present
    df.insert(loc=df.columns.get_loc('cust_name_id'), column='Handling', value=['' for i in range(df.shape[0])])
    df['Handling'] = np.where(df.section=='SRVCHG', 20, None)
    
    # clear out service charge rows - leaving acct_id (for grouping), description and name (don't want to lose any info)
    small_df = df.drop(columns=['Handling', 'acct_id', 'description', 'name'])
    small_df.loc[small_df['section'] == 'SRVCHG'] = None
    df[small_df.columns] = small_df
    
    # group by acct_id as a set
    df = df.groupby('acct_id', as_index=False).agg(lambda x: list(x))
    
    return df


def condense(original_data: pd.DataFrame):
    """Given a dataframe with nearly duplicate rows, condenses these into one row with
    multiple columns. 

    Args:
        original_data (pd.DataFrame): DataFrame from importing the input excel sheet

    Returns:
        pd.DataFrame: the condensed DataFrame
    """
    new_df = pd.DataFrame(original_data['acct_id'])
    list_cols = ['first_seat', 'last_seat', 'row_name', 'section', 'owed_seat', 'price_code', 'cost_per_seat', 'total_cost']

    # rename owed_amount column to owed_seat
    if 'owed_amount' in original_data.columns:
        original_data.rename(columns={'owed_amount': 'owed_seat'}, inplace=True)

    for col in original_data:
        # already have acct_id column - skip it
        if col == 'acct_id':
            continue
        # remove duplicate entries from all columns except those in list_cols
        elif col not in list_cols:
            set_vals = original_data[col].apply(lambda x: set([y for y in x if pd.notna(y)]))
        else:
            set_vals = original_data[col].apply(lambda x: [y for y in x if pd.notna(y)])
        expanded_col = pd.DataFrame(set_vals.values.tolist())
        
        # if there are multiple columns, rename them to [col_name]# format    
        if len(expanded_col.columns) > 1:
            expanded_col.rename(columns=lambda x: col+str(x+1), inplace=True)
        else:
            expanded_col.columns = [col]
            
        # put dates into m/d/YYYY format for Excel
        if col == 'renewal_date':
            expanded_col['renewal_date'] = expanded_col['renewal_date'].dt.strftime('%m/%d/%Y')
        new_df = pd.concat([new_df, expanded_col], axis=1)
        
    # add total_ticket_cost column in front of Handling
    new_df.insert(new_df.columns.get_loc('Handling'), 'total_ticket_cost', new_df[list(new_df.filter(regex='owed_seat'))].sum(axis=1))
    
    # add total due column = total ticket_cost + handling in front of cust_name_id
    new_df.insert(new_df.columns.get_loc('cust_name_id'), 'total_due', new_df['total_ticket_cost'] + new_df['Handling'].fillna(0))
    
    # reorder columns so that first_seat1, first_seat2, ..., last_seat1, last_seat2, ... is
    # first_seat1, last_seat1, first_seat2, last_seat2, ...
    column_order_idx = reorder_cols_alternate_seats(new_df)
    new_df = new_df.iloc[:, np.array(column_order_idx)]
    
    return new_df


def write_output(filepath, sheet_name, data):
    """Writes a dataframe to an Excel worksheet.

    Args:
        filepath (string): absolute file path to the location at which to save the excel sheet
        sheet_name (string): the desired name of the worksheet
        data (pd.DataFrame): the DataFrame to write to the Excel sheet
    """
    writer = pd.ExcelWriter(filepath)
    data.to_excel(writer, sheet_name=sheet_name, index=False)
    
    # get column indices for the money types
    money_col = ['total_ticket_cost', 'Handling', 'total_due'] + [col for col in data.columns if 'owed_seat' in col]
    money_idx = [data.columns.get_loc(col) for col in money_col]
    money_fmt = writer.book.add_format({'num_format': '$#,##0.00'})
    
    # auto-size columns
    for column in data:
        column_length = max(data[column].astype(str).map(len).max(), len(column))
        col_idx = data.columns.get_loc(column)
        # if error on set_column: pip install xlsxwriter
        
        # write money format for above money_cols
        if col_idx in money_idx:
            writer.sheets[sheet_name].set_column(col_idx, col_idx, column_length, money_fmt)
        else:
            writer.sheets[sheet_name].set_column(col_idx, col_idx, column_length)
 
    writer.close()


def reorder_cols_alternate_seats(df: pd.DataFrame):
    """Given a dataframe, returns the column indices reordered such that the seat information is 
    alternated as follows: first_seat1, last_seat1, first_seat2, last_seat2, ...
    All other columns remain in place.

    Args:
        df (pd.DataFrame): the dataframe containing the columns to reorder

    Returns:
        [int]: the new order the columns should be in. Note that the df is not changed.
    """
    reordered = []    
    col_names = [col for col in df.columns if 'first_seat' in col or 'last_seat' in col]
    if len(col_names) > 2:
        idx_last = col_names.index('last_seat1')
        idx_last_counter = idx_last
        for i in range(0, idx_last):
            reordered.append(col_names[i])
            if idx_last_counter < len(col_names):
                reordered.append(col_names[idx_last_counter])
                idx_last_counter += 1
    else:
        reordered = reordered + col_names
    
    reordered_idx = [df.columns.get_loc(col) for col in reordered]
    last_idx = np.array(reordered_idx).max()
        
    new_idx_order = [i for i in range(0,reordered_idx[0])] + reordered_idx + [i for i in range(last_idx+1, df.shape[1])]
    return new_idx_order


def append_sales_reps(df, sales_rep_file_path, output_diff):
    """Performs a left merge on df and the data read from sales_rep_file_path. This means all rows
    present in df will remain, and only those rows in sales_rep_file_path data that correspond to a row
    in df will be added. 

    Args:
        df (pandas.DataFrame): The dataframe to merge onto
        sales_rep_file_path (String): A file path to an Excel sheet containing the account rep information
        output_diff (boolean): True if the program should print the acct_ids not merged. False otherwise.

    Returns:
        pandas.DataFrame: The dataframe corresponding to df + the columns from sales_rep_file_path
    """
    # open sales rep file and read
    sr_df = pd.read_excel(sales_rep_file_path)
    
    appended = df.merge(sr_df, how='left', on='acct_id')
    appended['add_date'] = appended['add_date'].dt.strftime('%m/%d/%Y')
    
    if output_diff:
        merged = df.merge(sr_df, how='outer', indicator=True)
        print('Acct ID present only in condensed:')
        print(merged[merged['_merge'] == 'left_only']['acct_id'].tolist())
        print()
        print('Acct ID present only in sales rep:')
        print(merged[merged['_merge'] == 'right_only']['acct_id'].tolist())
    return df


def parse_args():
    """Reads and stores command-line arguments.

    Returns:
        Namespace: the command-line arguments that were passed to the program
    """
    parser = argparse.ArgumentParser(description='main.py')
    parser.add_argument('--input_file', dest='input_file', required=True, type=str, help='file path for input excel file.')
    parser.add_argument('--output_file', dest='output_file', required=True, type=str, help='file path for output excel file.')
    parser.add_argument('--sheet_name', dest='sheet_name', required=False, default='Summary', help="The desired name of the worksheet for the outputted excel file.")
    parser.add_argument('--sales_rep_file', dest='sales_rep_file', required=False, help="file path for the excel file containing sales rep information to append to the output excel file.")
    args = parser.parse_args()
    return args 


if __name__ == "__main__":
    args = parse_args()
    original_data = read_data_and_group(args.input_file)
    condensed_data = condense(original_data)
    
    if args.sales_rep_file is not None:
        condensed_data = append_sales_reps(condensed_data, args.sales_rep_file, True)
        
    write_output(args.output_file, args.sheet_name, condensed_data)