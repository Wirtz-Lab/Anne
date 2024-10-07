#Evaluating Skin Model Performance

import task1
import task2
import task3
import task4

source_file = '.././data/model_v165_input.xlsx'

# Task 1: Sum matrix
sum_train_path, sum_test_path = task1.sum_matrix(source_file)
print("\nSum matrices successfully created and saved.")

# Task 2: Merged matrix
merged_train_path = task2.merge_matrix(sum_train_path, False) #isTrain = False
merged_test_path = task2.merge_matrix(sum_test_path, True) #isTrain = True
print("\nMerged matrices successfully created and saved.")

# Task 3: Percent Matrix
percentage_train_path = task3.percent_matrix(merged_train_path, False) # isTrain = False
percentage_test_path = task3.percent_matrix(merged_test_path, True) # isTrain = True
print("\nPercentage matrices successfully created and saved.")

# Task 4: Precision Matrix
task4.precision_matrix(percentage_train_path, percentage_test_path)
print("\nPrecision matrix successfully created and saved.")

# All output files saved in the folder
