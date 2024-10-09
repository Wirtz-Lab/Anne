#Evaluating Skin Model Performance

import os, glob
import task1
import task2
import task3
import task4

dltlid = 92

src = r'\\10.99.68.54\Digital pathology image lib\HubMap Skin TMC project\240418_DLTL_master\DLTL_v{dltlid:d}\TrainCNN MDL\performance metrics'.format(dltlid = dltlid)
source_file = glob.glob(os.path.join(src,'net_*-01_trainingConfusionMetric.xlsx'))[0];
dst = os.path.join(src,'metrics')
if not os.path.exists(dst): os.mkdir(dst)

# Task 1: Sum matrix
sum_train_path, sum_test_path = task1.sum_matrix(source_file, dst)
print("\nSum matrices successfully created and saved.")

# Task 2: Merged matrix
merged_train_path = task2.merge_matrix(sum_train_path, False, dst) #isTrain = False
merged_test_path = task2.merge_matrix(sum_test_path, True, dst) #isTrain = True
print("\nMerged matrices successfully created and saved.")

# Task 3: Percent Matrix
percentage_train_path = task3.percent_matrix(merged_train_path, False,dst) # isTrain = False
percentage_test_path = task3.percent_matrix(merged_test_path, True,dst) # isTrain = True
print("\nPercentage matrices successfully created and saved.")

# Task 4: Precision Matrix
task4.precision_matrix(percentage_train_path, percentage_test_path,dst)
print("\nPrecision matrix successfully created and saved.")

# All output files saved in the folder
