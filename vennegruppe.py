#################################################################
#                                                               #
#     The user should use the provided template excel file.     #
#                                                               #
#     !! IMPORTANT, the archief sheet (Arkiv) should contain    #
#     all the children currently in the class. This list will   #
#     be used to make the new friend groups                     #
#                                                               #
#################################################################


import math
import sys
import random
from openpyxl import load_workbook


#################################################################
#                                                               #
#                   Checking the infile                         #
#                                                               #
#################################################################


# Check for right number of arguments
if len(sys.argv) != 2:
    print("Usage: python vennegruppe.py infile")
    sys.exit(1)


# Open infile
print("Opening infile")
infile = sys.argv[1]
wb_infile = load_workbook(filename=infile, read_only=True)


# Check if the provided set-up excel file is used by checking sheet names
sheet1, sheet2, sheet_hidden = wb_infile.sheetnames
if not sheet1 == 'Vennegruppe' or not sheet2 == 'Arkiv':
    print("Please use the provided example file")
    sys.exit(1)


#################################################################
#                                                               #
#           Read and store the previous groups                  #
#                                                               #
#################################################################


# grab the groups worksheet (Vennegruppe)
sheet1_infile = wb_infile["Vennegruppe"]


# Initialize dictionary to hold the previous friendgroups (Vennegruppe)
groups = {}
number_of_groups = 0
number_of_kids_in_group = 0


# grab the previous friendgroups (Vennegruppe)
print("\nReading latest friendgroups")
while True:
    number_of_groups += 1

    if not sheet1_infile.cell(row = 1, column = number_of_groups).value:
        break

    groups[number_of_groups] = list()

    # Grab the names of the children in the group
    while True:
        number_of_kids_in_group += 1

        if not sheet1_infile.cell(row = (number_of_kids_in_group + 1), column = number_of_groups).value:
            number_of_kids_in_group = 0
            break

        groups[number_of_groups].append(sheet1_infile.cell(row = (number_of_kids_in_group + 1), column = number_of_groups).value)
        

#################################################################
#                                                               #
#   Store history of 'who_has_been_in_a_group_with' in a dict   #
#                                                               #
#################################################################


# Change to archive sheet (Arkiv)
sheet2_infile = wb_infile["Arkiv"]


# Initialize variables to store 'who_has_been_in_a_group_with'
archief = {}
child = 0
child_name = ""
has_been_in_group_with = list()
iterator = 3


# grab the name and gender of all the children (Kjønn, Navn)
print("Storing all the previous friend groups")
while True:
    child += 1

    if not sheet2_infile.cell(row = (child + 1), column = 1).value:
        break
    
    child_name = sheet2_infile.cell(row = (child + 1), column = 1).value

    # grab all the names of "has_been_in_group_with" (Har vært i gruppen med)
    while True:
        if not sheet2_infile.cell(row = (child + 1), column = iterator).value:
            iterator = 3
            break
        
        has_been_in_group_with.append(sheet2_infile.cell(row = (child + 1), column = iterator).value)
        iterator += 1
    
    archief[child_name] = {"gender": sheet2_infile.cell(row = (child + 1), column = 2).value, "has_been_in_group_with": has_been_in_group_with}
    has_been_in_group_with = list()
    

#################################################################
#                                                               #
#       Add lastest groups to the "has_been_in_group_with"      #
#                                                               #
#################################################################


# Access each group from the latest groups
print("Adding the latest groups to the archief")
for group in groups:
    
    #Access the names in the group
    for name in groups[group]:

        # Check if name is in the archief
        if name in archief.keys():
        
            # Add the other children of the latest group into the archief
            for kid in groups[group]:
                if name != kid:
                    archief[name]["has_been_in_group_with"].append(kid)
        

#################################################################
#                                                               #
#                   Make new random groups                      #
#                                                               #
#################################################################


# Make a list containing all the names
print("\nCreating new friend groups")
randomized_list_of_names = list(archief.keys())


# Function to randomize the order of the names in the list
def randomizeChildren():
    random.shuffle(randomized_list_of_names)
    return randomized_list_of_names


# Initializing variables to make new friend groups
total_nr_of_groups = 1
nr_of_5_kids_per_group = len(archief)%4
nr_of_4_kids_per_group = int((len(archief) - nr_of_5_kids_per_group * 5) / 4)
new_friend_groups = {}
genders_per_group = {}

# Calculate how many groups are needed
length_archief = len(archief)
if length_archief >= 15:
    total_nr_of_groups = math.floor(len(archief)/4)
else:
    if length_archief % 4 <= 2:
        total_nr_of_groups = math.floor(len(archief)/4)
    else:
        total_nr_of_groups = math.ceil(len(archief)/4)



# Function to make new groups from the randomized list
def makeGroups():
    randomizeChildren()
    iterator = 1

    # Make the groups with 5 children
    for group_with_five in range(0, (nr_of_5_kids_per_group * 5), 5):
        # Assigns the friend group number and stores the names
        new_friend_groups[iterator] = randomized_list_of_names[group_with_five:(group_with_five + 5)]
        iterator += 1
    
    # Make the groups with 4 children
    for group_with_four in range((nr_of_5_kids_per_group * 5), (nr_of_4_kids_per_group * 4 + nr_of_5_kids_per_group * 5), 4):
        # Assigns the friend group number and stores the names
        new_friend_groups[iterator] = randomized_list_of_names[group_with_four:(group_with_four + 4)]
        iterator += 1

# Function to make a list of all the genders in one group
def genders():
    makeGroups()
    genders_per_group = {}
    tmp_genders = list() # empty list to hold the genders of the children in one group


# Function to score the new friend group
#    5 points for same gender
#   10 points for having been in same group before
def scoring():
    iterator = 1
    for group in new_friend_groups:
        print("\nVennegruppe", iterator, end=":  ")
        iterator += 1
        for child in new_friend_groups[group]:
            print(child, end=", ")


makeGroups()
print("Total number of children: {0}\nTotal number of groups: {1} \n {2} groups with 5 children and," 
      "\n {3} groups with 4 children".format(length_archief, nr_of_5_kids_per_group + nr_of_4_kids_per_group,
      nr_of_5_kids_per_group, nr_of_4_kids_per_group))

print("\n", genders_per_group)

# Print the new friends group to the screen
iterator = 1
for group in new_friend_groups:
    print("\nVennegruppe", iterator, end=":  ")
    iterator += 1
    for child in new_friend_groups[group]:
        print(child, end=", ")


# Data can be assigned directly to cells
#x = random.randint(1, 10)
#ws['A1'] = x

# Save the file
#wb.save("sample.xlsx")

# Empty dictionaries
groups = {}
archief = {}

# Success
sys.exit(0)

#################################################################
#                                                               #
#               Written by Hanna Noordzij 2019                  #
#                                                               #
#################################################################