import random
from openpyxl import load_workbook

mice_input = []
mice_weights_input = []

#INPUTS
#Name of the excel file before the .xlsx extension in inverted commas
#NOTE: The excel file must be in the same folder as this script
file_name = "Mice weights"
#Number of groups that need to be made
no_of_groups = 5
#Size of each Group
group_size = 5
#Maximum increment and decrement of group weight from average weight of a group
difference = 0.45
#Max number of times it tries a certain set of numbers
max_attempts=20000
#You can play around with these two settings (max_attempts and difference) for more precise grouping or if code is taking too long to run 
#Excel Spreadsheet Information
column_with_mice_names = "A"
column_with_mice_weights = "B"
row_with_first_value = 2
row_with_last_value = 26

#Code
book = load_workbook(f"{file_name}.xlsx")
sheet = book.active
for i in range(row_with_first_value, row_with_last_value + 1):
  mice_input.append(sheet[f"{column_with_mice_names}{i}"].value)
  mice_weights_input.append(sheet[f"{column_with_mice_weights}{i}"].value)

y = len(mice_input) - 1
values_of_n = [y]
while len(values_of_n) < no_of_groups:
  y -= group_size
  values_of_n.append(y)

total_mice_weight = sum(mice_weights_input)
ideal_group_weight = total_mice_weight * group_size / len(mice_input)
print(f"Ideal Weight of each group should be {ideal_group_weight}"+"\n")

program_flag = False
def create_groups():
  global program_flag
  mice = mice_input.copy()
  mice_weights = mice_weights_input.copy()
  group_weights = []
  group_mice = []
  attempts = 0
  function_flag = True
  for i in range(no_of_groups - 1):
    status = False
    if function_flag == False:
      return
    while status == False:
      if attempts == max_attempts:
        function_flag = False
        break
      attempts += 1
      mice_copy = mice.copy()
      mice_weights_copy = mice_weights.copy()
      n = values_of_n[i]
      group_weight = 0
      picked_mice = []
      mice_names = []
      while len(picked_mice) < group_size:
        x = random.randint(0, n)
        if x not in picked_mice:
          picked_mice.append(x)
          mice_names.append(mice_copy[x])
          group_weight += mice_weights_copy[x]
          del mice_copy[x]
          del mice_weights_copy[x]
          n -= 1
      if ideal_group_weight - difference < group_weight < ideal_group_weight + difference:
        status = True
        group_mice.append(mice_names)
        group_weights.append(group_weight)
        for i in picked_mice:
          del mice[i]
          del mice_weights[i]
  group_weight = 0
  mice_names = []
  for i in range(len(mice)):
    mice_names.append(mice[i])
    group_weight += mice_weights[i]
  group_weights.append(group_weight)
  group_mice.append(mice_names)
  
  if ideal_group_weight - difference < group_weight < ideal_group_weight + difference:
    print(f"Group Names: {group_mice}")
    print(f"Group Weights: {group_weights}")
    print("\n"+f"The difference between the heaviest group and the lightest group is {max(group_weights)-min(group_weights)}")
    program_flag = True

while program_flag == False:
  create_groups()
