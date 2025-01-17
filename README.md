# Welcome!

This is a professor email generator which takes in a draft email and a list of professors in a single excel file and generates the professor emails within the excel file. You can then copy and past the emails and the respective email adresses into your bxscience account to send them. (I might make a program that allows you to auto send a list of emails after you review them in excel, but that's for later)

There is an [alternate version](https://github.com/JosephOziel/Automated_Emails) by one of the period 3 TA's, but I would advise against using it as it doesn't let you preview any of the emails prior to sending. Also if you screw something up, then you can't fix it before sending as the generation and sending are done all in one go. ***If for some reason you'd like to risk your emails and use that program, don't come complaining to me that something went wrong. Go to one of the period 3 TA's and complain to them.***

## How to Use:

### Downloads
First download the following files (and ideally put them in a seperate folder to avoid having issues with duplicate names in your directory):

- [Excel Sheet](docs/PROF_EMAIL_LIST.xlsx)
- [The Python Program](docs/Generator.py)

### The Excel File

Open the excel file and you should see two sheets at the bottom (if you're using excel, if not then there should be two tabs) one that says INPUT and one that says OUTPUT. The INPUT sheet is where you'll put all of your information, the professor email, name, and your email draft in the respective boxes. The OUTPUT sheet is where you'll look to get your final results to then copy and paste.

<p align="center">
  <img width="234" alt="Screenshot 2025-01-17 at 8 27 31 AM" src="https://github.com/user-attachments/assets/60d46a21-6690-4743-a95d-70f04e1ad5ae" />
</p>

Now look in the INPUT sheet and you should see three columns, one that says Prof Full Name, one that says the email, and the last saying the email draft.

<p align="center">
  <img width="294" alt="Screenshot 2025-01-17 at 8 29 50 AM" src="https://github.com/user-attachments/assets/bc083817-a677-43c7-9683-0abf3f9b783a" />
</p>

I think it's self explanatory but just add the professor's full name (note that only the last name will be used, there are ways to change the code but I'll mention that in [settings]), and email where it says email. Note that the code stops bringing in data the moment there is an empty cell below the headings, so make sure there are none so that the program gets all the professors you need.
As for the email draft, simply copy and paste the draft, but it has to fit in the cell only. If you only copy and paste from google docs, then you'll get the email spread over many cells, which won't generate the proper email. To do this properly, copy the email you want to use then paste it in the cell value section instead of a simply copy and paste.

<p align="center">
  <img width="1678" alt="Screenshot 2025-01-17 at 8 35 29 AM" src="https://github.com/user-attachments/assets/ec6b4378-909f-41a6-ae1e-6adac96c7111" />
</p>

(I'll mention how to do all of this in [google sheets] in a bit, its basically the same but you have to donwload the file at the end and [if you have one] place it in the folder with the program)

Once you have all the information in, you are now ready to start the program!

### The Program

Now the program `Generator.py` has a handful of settings that you may need to adjust prior to running:

- `workbook_name` This is the name of the excel sheet that will have all the inputs and outputs of the program, its by default the same one here so unless something went wrong with naming the file or you changed it you shouldn't need to change this.
- `

