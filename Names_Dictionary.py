#This program will initialize a dictionary with names.

#Initialize a dictionary of last names.
last_names={
    "Marni": "Hager",
    "Joey": "Marianer"
}

#Print desired information.
print(last_names)
print(last_names["Joey"])

#Change the last name of Joey in the dictionary.
last_names["Joey"]="Galston"

#Validate the name change.
print(last_names["Joey"])

#Add new names to the dictionary.
last_names["Angela"]="Brett"

#Validate the name was appended.
print(last_names)

#Ask the user for a first name.
Requested_Name=input("Please enter a first name: ")

#If name exists, print the last name.  Otherwise,
#inform user the name does not exist.
if Requested_Name in last_names:
    print(last_names[Requested_Name])
else:
    print("That name does not exist.")