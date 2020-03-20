# School-Hydration

These set of scripts are designed to help your school quickly get up and running using Microsoft Teams. There is one script for creating Office 365 accounts for students and teachers. Another script for creating a Teams site for each class and assigning the teacher to it, and a final script for assigning the students to their classes.

I realize that not everyone who is in need of the solution, is a PowerShell guru. So I have provided detailed steps below on how you can implement this for your school.

# Requirements
- Global admin rights to the Office 365 tenant
- [AzureAD PowerShell Module](https://www.powershellgallery.com/packages/AzureAD/2.0.2.4)
- [MicrosoftTeams PowerShell Module](https://www.powershellgallery.com/packages/MicrosoftTeams/1.0.5)

# Importing Students and Teachers

Students and teachers are both imported using the same script. If you already have Office 365 accounts for your students and teachers you can skip this step and go to the class creation scripts. If you are assigning different licenses to teacher and students, it is recommended that you run it once for teachers and once for students. See the sample CSV file Import-Teachers.csv and Import-Students.csv for examples of how to format your data. You only need to enter their first and last name and if they are a teacher, student, or faculty. Once you've filled out the CSV run the **Import-StudentsAndTeachers.ps1** script to begin the import process. The usernames will be auto-generated to be the person's firstname.lastname and any duplicates will have a number appended to the end of them.

When the initialize the script you will be prompted to enter your credentials for your Office 365 tenant, to choose the email domain, and the licenses you want to assign. When the script completes it will output a second CSV file in the same path as the orginal with "-imported" appended to the file name. This file will contain the username for each account created, that you will need to populate the classes and assignments CSV files.

# Create Team sites for Classes

Once you have your user accounts created, you can now create a Team site for each class. All you need to do is provide a CSV file with the name of the class and the teacher. Each class name must be unqiue and the teacher value should be the UserPrincipalName that was created by the import process and found in the CSV it exported. See the CSV Create-ClassTeams.csv for an example. 

To import your classes simply run the **Create-ClassTeams.ps1**, passing in the path to your CSV file. One you start the script you will be prompted to enter your credentials, then it will start creating the Teams sites.

# Create Team sites for Classes

Now that you have your user accounts created and your Team sites setup, you can now assign students to thier class. To do this you need a CSV file with the name of the class and the UserPrincipalName of the student to assign to it. See the CSV Assign-Classes.csv for an example. 

To assign students to classes run the **Assign-Classes.ps1**, passing in the path to your CSV file. One you start the script you will be prompted to enter your credentials, then it will go through and assign all the student to their classes.

# Detailed Instructions

## Download Scripts
1. At the top of this page click the Clone of download button
2. Select Download Zip

![RepoDownload](Screenshots\RepoDownload.png)

3. When the download completed extract the zip file

## Install Required Modules
1. Open PowerShell console
2. Install the Azure AD module using the command:
```powershell
Install-Module AzureAD -Scope CurrentUser
```
3. If prompt to trust repository enter Y for Yes.
4. Install the Microsoft Teams module using the command:
```powershell
Install-Module MicrosoftTeams -Scope CurrentUser
```
5. Again, if prompt to trust repository enter Y for Yes.

## Import Teachers and Students
1. Open the Import-Teachers.csv and Import-Students.csv in Excel and enter the information for your school.
2. In the PowerShell console navigate the the folder you extracted the zipped files to.
3. Start the teacher import by running the command below. If your CSV is in the same folder as your script you can leave the path as .\\. Otherwise you will need to provide the full path.
```powershell
.\Import-StudentsAndTeachers.ps1 -csvPath .\Import-Teachers.csv -defaultPassword 'AStrongPassword'
```

![Import-Teachers](Screenshots\Import-Teachers.PNG)

4. You will see a prompt to enter your username and password for Office 365. Enter your username and password for your global admin account.

![login](Screenshots\login.PNG)

5. Next you will be prompted to select the email domain for the users. Highlight the one you want to use and click OK

![domain](Screenshots\domain.PNG)

6. Next you will be prompted to select the license to assign to the users. Highlight the one you want to use and click OK

![license](Screenshots\license.PNG)

7. Once the import completes it will tell you where the exported CSV file is. You will need this later for creating the classes and assigning the students.

![Import-Teachers-done](Screenshots\Import-Teachers-done.PNG)

8. Repeat the process above for the Import-Students.csv file.

## Create Class Teams

1. Open the Create-ClassTeams.csv in Excel and enter the information for your school. The class names should all be unique and the value in the teacher column should be the UserPrincipalName found in the export file from the teacher import. 
2. In the PowerShell console navigate the the folder you extracted the zipped files to.
3. Start the class creation by running the command below. If your CSV is in the same folder as your script you can leave the path as .\\. Otherwise you will need to provide the full path.
```powershell
.\Create-ClassTeams.ps1 -csvPath .\Create-ClassTeams.csv
```

4. You will see a prompt to enter your username and password for Office 365. Enter your username and password for your global admin account.

![login](Screenshots\login.PNG)

5. You will see a list of the Team sites as they are created. The teacher will automatically be added as the Owner of the Team.


![Create-ClassTeams](Screenshots\Create-ClassTeams.PNG)

## Assign Students to Classes

1. Open the Assign-Classes.csv in Excel and enter the information for your school. The class names should match the names from the Create-ClassTeams.csv and the value in the student column should be the UserPrincipalName found in the export file from the student import. 
2. In the PowerShell console navigate the the folder you extracted the zipped files to.
3. Start the class assignment below by running the command below. If your CSV is in the same folder as your script you can leave the path as .\\. Otherwise you will need to provide the full path.
```powershell
.\Assign-Classes.ps1 -csvPath .\Assign-Classes.csv
```

4. You will see a prompt to enter your username and password for Office 365. Enter your username and password for your global admin account.

![login](Screenshots\login.PNG)

5. Once the script completes you are finished.

![Assign-Classes](Screenshots\Assign-Classes.png)