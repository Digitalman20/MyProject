#**Employee Management**




##Video Demo:



##***Project.py file***



###**_Descrition_**:
The project basically stores the records of employees in a file in a small company. We could perform tasks such as 'hire','fire', 'change hourly pay' and finally 'mail' the file.


###***Working and Use of each function***

####***Function basic_details()***:

#####***Working***:
This function uses a while loop to take input. A variable 'add_more' is used as a condition for the while loop. The while loop first prompts for the name, designation, and hourly pay of the employee that needs to be added and makes a dictionary of these 3 keys, and then the dictionary is appended to a list(details). At the end of the loop, I added a prompt to add one more person. If the input for the prompt is 'y' we can add one more person, anything else we exit the loop. After exiting this function returns a list of dictionaries. The input for designation can only be the keys from the valid_post dictionary. The value of each key in the dictionary is a tuple containing the pay range. The pay can only be a value in the range. Initially, I figured catching an 'EOFError' to exit the loop was ideal but I changed it to the current design when I realized the 'extra' prompt for the name just sitting there before catching the error.

#####***Use***:
This function contributes nothing to the project 'alone' but when paired with functions like *xl()* and *hire()*, we can use this to get multiple inputs to store in an excel sheet.

####***Function xl(list, file)***:

#####***Working***:
This function takes in the return value of *basic_details()* or any list of dictionaries(with just 3 key-value pairs)and the file's name. First, a workbook is created with a worksheet in it. Then the headers(S.no,Name,Designation,Pay) are written in the first row using the cell method. The for loop reads through the list of dictionaries(details) and writes the contents of each dictionary in a new line. Basically, the loop appends the serial number, name, designation, and pay in a new line without overwriting the existing rows. Finally, the save method is used to update the file if it already exists or make a new one.

#####***Use***:
This function is used to write a new excel file or overwrite an existing file. This may seem quite odd about overwriting an existing file but it is useful to replace a bunch of existing people.

####***Function hire(list,file)***:

#####***Working***:
This function takes in the return value of *basic_details()* or any list of dictionaries(with just 3 key-value pairs) and the file's name. This function can only be applied if the mentioned file exists. First, an existing file is opened in the background. This function appends the serial number, name, designation, and pay to a new row at the end of the sheet using the for loop. This function writes the serial number based on the max_row method. Finally, the workbook is saved with the updates using the save method.

#####***Use***:
This function is basically used to hire(add or append) one or more person(s) to the company.

####***Funtion fire()***:

#####***Working***:
This function prompts for an input, the name of the person who needs to be fired. The existing excel file is opened in the background. Then, a nested for loop goes through each cell in column 'B' and checks for the input in the cells, if it finds a match it erases the row from the excel sheet. I've noticed a bug prior to adding the last if statement in the nested for loop. Basically, an empty row is formed at the end of excel sheet every time this function is applied, this might not seem like a big issue. As the *hire()* function uses the last row number to write the serial number of the next person, the serial number comes out different than expected and a row is left out in the middle of the excel which isn't visually pleasing. Finally saving the updated file.

#####***Use***:
This function is used to fire(remove from the file) a person from the company. This can only be used to remove one person at any given time.

####***Function pay_raise(to_raise, r, file)***:

#####***Working***:
This function takes in 3 parameters. 'to_raise' being the name of the person receiving a hike, 'r' the hike amount, and the name of the file. First, the existing excel file is opened, then a for loop iterates through the column 'B' and matches the name. If a match is found, the pay in the row is taken and the numeric part is stored in a variable and r is added to the variable finally a string is created in the required pattern and it is written in place of pay in that row.

#####***Use***:
This function is used to increase the hourly pay of a person in the company. This can only be used for one person at any given time

####***Function pay_cut(to_cut, c, file)***:

#####***Working***:
This function takes in 3 parameters.'to_cut' being the name of the person receiving a pay cut, 'c' the amount being cut, and the name of the file. First, the existing workbook opens in the background and the nested for loop goes through each cell in column 'B' and matches the name. If the match is found, the position of the hourly pay of that person is stored. The hourly pay at that stored position is taken and split from the '$' sign and the numeric part is turned to an int and the pay cut is subtracted from it. Then an if-else statement checks if the pay is less than or equal to zero a statement is output and the while loop in main takes care of the prompt. If it is greater than 0 a string is created and it overwrites the cell in the previously stored position and the workbook is saved.

#####***Use***:
This function is used to increase the hourly pay of a person in the company. This can only be used for one person at any given time.

####***Function mail()***:

#####***Working***:
The first thing this function does is decrypt the password of the mail address that I made for this project. I encrypted the password using the cryptography module. Basically, a computer generated is written in a file(key.key) and the encrypted password is stored in another file(password key). While decrypting the key file is read and the content is stored in a variable called a key, similarly the encrypted password is stored in a variable. The key is important to decrypt without the key the decryption isn't possible. After decryption, the password is stored in a variable. Then the SMTP object connects to the server and the hardcoded mail address and the variable containing the password are used to log in. There is an input for the mail address. The message is generated using the email module. Then this message is sent.

#####***Use***:
This function is straightforward and sends the file to the input mail address. This is used in every menu option in the main.

####***Function mail_choice***:

#####***Working***:
This function has an input prompt that asks the user if they want to mail the file. If 'y' the mail goes through.

#####***Use***:
The sole purpose of this function is to execute the mailing process in every menu option except the 'mail' menu option.


###***main()***:
The while loop in the main works on the condition of do_it_again. In the loop, first, all the menu options are printed for the user. There is another while loop that takes input to choose the menu option, we get another prompt if the menu option is not valid.

####***Menu option 1***:
This menu option is given to write or overwrite an excel file. First the *basic_details* runs to get a list of dictionaries this is soon followed by the *xl* function. The *mail_choice* is used to ask the user if he/she wanted to mail the file, and finally, another prompt is asked if the user wanted, to do anything else.

####***Menu option 2***:
This menu option is given to hire one or more person(s). First the *basic_details* runs to get a list of dictionaries, this is followed by the *hire* function to add them to the file. The *mail_choice* is used to ask the user if he/she wanted to mail the updated file, and finally, another prompt is asked, if the user wanted to do anything else.

####***Menu option 3***:
This menu option is given to fire one person. First  *fire* runs to remove the person from the file.  Then *mail_choice* is used to ask the user if he/she wanted to mail the updated file, and finally, another prompt is asked if the user wanted to do anything else.

####***Menu option 4***:
This menu option is given to make changes to a person's pay. In this a while loop ensures a proper input value as in this menu option there are further more options. '6' for pay raise and '7' for a pay cut. If the user types '6' input for the name is asked and then input for the amount is asked, a while loop makes sure that the input for the amount is an integer, these inputs are used as the parameters for the *pay_raise*. If the user types '7' input for the name is asked and then input for the amount is asked, a while loop inside a while loop makes sure that the input for the amount is an integer, these inputs are used as the parameters for the *pay_cut*. Then if the value of the result(i.e. pay-amount) is greater than zero the loop breaks else the loop repeats. Then *mail_choice* is used to ask the user if he/she wanted to mail the updated file, and finally, another prompt is asked if the user wanted to do anything else.

####***Menu option 5***:
This menu option basically runs *mail* and a prompt is asked if the user wanted to do anything else.



##***test_project.py***:


###***test_pay_raise***:
This test has a list of dictionaries which is used as a parameter for the xl function to write a test_employee.xlsx file. The *pay_raise* function runs with to_raise, r, test_employee.xlsx as its parameters. This updates the pay of to_raise(an employee). The first test asserts the updated pay to the sum of current pay and r. The second test asserts a name not in the file to the output statement if the name isn't in the file.


###***test_pay_cut***:
This test has a list of dictionaries which is used as a parameter for the xl function to write a test_employee.xlsx file. The *pay_cut* function runs with to_cut, c, and test_employee.xlsx as its parameters. This updates the pay of to_cut(an employee). The first test asserts the updated pay to the difference between current pay and c. The second test asserts a negative or a zero result to the output statement for that condition. The last test asserts  a name not in the file to the output statement if the name isn't in the file.


###***test_xl***:
This test has a list of dictionaries which is used as a parameter for the xl function to write a test_employee.xlsx file. All the assertions in this assert the values in the cells to the values in the list.




