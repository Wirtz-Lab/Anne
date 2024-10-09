# TASK 1
import os

import pandas as pd

def sum_matrix(source_file, dst = '../../Outputs'):
    df_dict = pd.read_excel(source_file, sheet_name=None, index_col=0, header=3)

    sum_dict = {}
    test_sum_dict = {}
    test_sheets_list = ['im05', 'im12', 'im14', 'im23', 'im24', 'im39', 'im45', 'im46']

    for sheet_name, df in df_dict.items():
        data_range = df.iloc[0:24, 0:26]  # Extract only relevant data

        # Convert all values to numeric, coercing errors to NaN and fill NaN with 0
        data_range = data_range.apply(pd.to_numeric, errors='coerce').fillna(0)
        data_range = data_range.astype('int64')

        # Iterate over each cell in the sheet
        for col in data_range.columns:  # Get col-name
            for row in data_range.index:  # Get index-name
                cell_value = data_range.at[row, col]  

                if sheet_name in test_sheets_list:
                    if (col, row) not in test_sum_dict:
                        test_sum_dict[(col, row)] = 0
                    test_sum_dict[(col, row)] += cell_value
                else:
                    if (col, row) not in sum_dict:
                        sum_dict[(col, row)] = 0
                    sum_dict[(col, row)] += cell_value

    sum_df = pd.DataFrame(0, columns=df.columns, index=df.index, dtype='int64')
    test_sum_df = pd.DataFrame(0, columns=df.columns, index=df.index, dtype='int64')

    for (col, row), total in sum_dict.items():
        sum_df.at[row, col] = total 
    for (col, row), total in test_sum_dict.items():
        test_sum_df.at[row, col] = total

    # # Display DFs
    # print("\nFinal Summed DataFrame:")
    # print(sum_df)
    # print("\nFinal Train Summed DataFrame:")
    # print(train_sum_df)

    # # Excel Formatting / Save DFs
    if not os.path.exists(dst): os.mkdir(dst)
    train_output_file_path = os.path.join(dst,'Sum_matrix_train.xlsx')
    with pd.ExcelWriter(train_output_file_path, engine='openpyxl') as writer:
        sum_df.to_excel(writer, sheet_name='trainSetSum', index=True, header=True)

    test_output_file_path = os.path.join(dst, 'Sum_matrix_test.xlsx')
    with pd.ExcelWriter(test_output_file_path, engine='openpyxl') as writer:
        test_sum_df.to_excel(writer, sheet_name='testSetSum', index=True, header=True)

    return train_output_file_path, test_output_file_path # Return the file paths