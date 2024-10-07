import task1
import task2
import task3
import task4

source_file = './model_v165_input.xlsx'

# Task 1: Sum matrix
sum_train_path, sum_test_path = task1.sum_matrix(source_file)

# Task 2: Merged matrix
merged_train_path = task2.merge_matrix(sum_train_path, False) #isTrain = False
merged_test_path = task2.merge_matrix(sum_test_path, True) #isTrain = True


# Task 3: Percent Matrix
percentage_train_path = task3.percent_matrix(merged_train_path, False) # isTrain = False
percentage_test_path = task3.percent_matrix(merged_test_path, True) # isTrain = True

# Task 4: Precision Matrix
task4.precision_matrix(percentage_train_path, percentage_test_path)

# All files saved in the folder