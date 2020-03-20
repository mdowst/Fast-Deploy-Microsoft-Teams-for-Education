# School-Hydration

These set of scripts are designed to help your school quickly get up and running using Microsoft Teams. There is one script for creating Office 365 accounts for students and teachers. Another script for creating a Teams site for each class and assigning the teacher to it, and a final script for assigning the students to their classes.

# Requirements
- Global admin rights to the Office 365 tenant
- [AzureAD PowerShell Module](https://www.powershellgallery.com/packages/AzureAD/2.0.2.4)
- [MicrosoftTeams PowerShell Module](https://www.powershellgallery.com/packages/MicrosoftTeams/1.0.5)

# Importing Students and Teachers

Students and teachers are both imported using the same script. If you already have Office 365 accounts for your students and teachers you can skip this step and go to the class creation scripts. If you are assigning different licenses to teacher and students, it is recommended that you run it once for teachers and once for students. See the sample CSV file Import-Teachers.csv and Import-Students.csv for examples of how to format your data. You only need to enter their first and last name and if they are a teacher, student, or faculty. Once you've filled out the CSV run the *Import-StudentsAndTeachers.ps1* script to begin the import process. The usernames will be auto-generated to be the person's firstname.lastname and any duplicates will have a number appended to the end of them.

When the initialize the script you will be prompted to enter your credentials for your Office 365 tenant, to choose the email domain, and the licenses you want to assign. When the script completes it will output a second CSV file in the same path as the orginal with "-imported" appended to the file name. This file will contain the username for each account created, that you will need to populate the classes and assignments CSV files.

# Create Team sites for Classes

Once you have your user accounts created, you can now create a Team site for each class. All you need to do is provide a CSV file with the name of the class and the teacher. Each class name must be unqiue and the teacher value should be the UserPrincipalName that was created by the import process and found in the CSV it exported. See the CSV Create-ClassTeams.csv for an example. 

To import your classes simply run the *Create-ClassTeams.ps1*, passing in the path to your CSV file. One you start the script you will be prompted to enter your credentials, then it will start creating the Teams sites.

# Create Team sites for Classes

Now that you have your user accounts created and your Team sites setup, you can now assign students to thier class. To do this you need a CSV file with the name of the class and the UserPrincipalName of the student to assign to it. See the CSV Assign-Classes.csv for an example. 

To assign students to classes run the *Assign-Classes.ps1*, passing in the path to your CSV file. One you start the script you will be prompted to enter your credentials, then it will go through and assign all the student to their classes.