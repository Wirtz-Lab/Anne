import os, glob
import pandas as pd
import numpy as np
def sum_matrix(source_file):
    df_dict = pd.read_excel(source_file, sheet_name=None, index_col=0, header=3)
    try:
        df_dict.pop('Sheet2')
    except:
        print('')
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
    return sum_df, test_sum_df # Return the file paths

def sort_sum_by_toi(df1, toi):
    newidx = [_ for _ in df1.keys().to_list() if _ not in toi] + toi
    # send toi to edge
    a = df1.loc[newidx, :]
    b = a.loc[:, newidx]
    # sum rows of toi and assign to first class of toi
    b.loc[toi[0], :] = b.loc[toi, :].sum(axis=0)
    # delete merged rows
    b = b.drop(toi[1:])
    # column
    b.loc[:, toi[0]] = b.loc[:, toi].sum(axis=1)
    # delete merged columns
    df2 = b.drop(toi[1:], axis=1)
    return df2

def merge_classes(df):
    merged_df = df.copy()  # cp the dataframe
    toi = ['RBC', 'endothelial', 'vessel_wall']
    merged_df = sort_sum_by_toi(merged_df, toi)
    toi = ['background', 'noise', 'shade', 'unknown']
    merged_df = sort_sum_by_toi(merged_df, toi)
    toi = ['ecm', 'inflammation', 'large_interstitial', 'papillary_dermis', 'reticular_dermis']
    merged_df = sort_sum_by_toi(merged_df, toi)
    return merged_df

def percent_matrix(df):
    row_sum = df.sum(axis = 1)
    col_sum = df.sum(axis = 0)
    return df.div(row_sum, axis=0), df.div(col_sum, axis=1)

def diagonal(df):
    return pd.Series(np.diag(df), index=[df.index])

def confusion_matrix_2_metrics(source_file):
    train_sum, test_sum = sum_matrix(source_file)
    train_merged = merge_classes(train_sum)
    test_merged = merge_classes(test_sum)
    train_precision_m, train_recall_m = percent_matrix(train_merged)
    test_precision_m, test_recall_m = percent_matrix(test_merged)
    train_precision = diagonal(train_precision_m)
    train_recall = diagonal(train_recall_m)
    test_precision = diagonal(test_precision_m)
    test_recall = diagonal(test_recall_m)

    return [train_precision.replace(0, np.nan).mean().round(5).item() * 100,
    train_recall.replace(0, np.nan).mean().round(5).item() * 100,
    test_precision.replace(0, np.nan).mean().round(5).item() * 100,
    test_recall.replace(0, np.nan).mean().round(5).item() * 100]

if __name__ == "__main__":
    dltlids = [42, 228]
    metrics = []
    for dltlid in dltlids: #iterate deep learning model versions
        try:
            # define input and output path
            src = r'\\10.99.68.54\Digital pathology image lib\HubMap Skin TMC project\240418_DLTL_master\DLTL_v{dltlid:d}\TrainCNN MDL\performance metrics'.format(
                dltlid=dltlid)
            source_file = glob.glob(os.path.join(src, 'net_*-01_trainingConfusionMetric.xlsx'))[0]
            # execute the conversion
            metric = confusion_matrix_2_metrics(source_file)
            metrics.append([dltlid]+metric)
        except:
            metrics.append([dltlid, 0, 0, 0, 0])
        # save as a csv for summary
        dst = os.path.join(src, 'output')
        if not os.path.exists(dst): os.mkdir(dst)
        pd.DataFrame(metrics).to_csv(os.path.join(dst,'tmp.csv'), index=False)