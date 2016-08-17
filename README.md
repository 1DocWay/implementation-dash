# iDash

A command line tool that generates the Implementation Dashboard spreadsheet from projects and tasks in Asana.


# Installation

Check to see if you have Pip installed on your system. You can test this by running:

    $ pip -V

If the command is not recognized, you do not have Pip. Pip generally comes bundled with a Python install. To install Python, including Pip and other setup tools: https://www.python.org/downloads/. **Install Python version 2.7.12**.

You may need to upgrade your Pip version. **Please use Pip version 8.1.2**. To upgrade, run the command:

    $ pip install --upgrade pip

Once Pip has been installed, clone this repository to your desktop, then navigate to the root directory. 

Simply run:

    $ pip install .
    
This command will find and install all the Python packages needed for the script to run. 

# Usage

### Create a Personal Access Token in Asana

You will need a Personal Access Token to run the script. This token serves as your authentication to retrieve projects and tasks from the Asana API. To create a token:

1. Log into Asana
2. Navigate to "My Profile Settings", and then switch over to the tab "Apps".
3. Click on "Manage Developer Apps"
4. You will see a section called "Personal Access Tokens". Click on "Create a New Personal Access Token".
5. Fill in a description for the token, and hit "Create". You will be given a token to copy and paste. Note that you will not be shown this token again. 
6. Copy and paste the token into a new text file. Save this file in the root directory of this project. Make sure to use the file extension ".txt".

### Running the script

You can run the script via the following command:

    $ idash TOKEN_FILE [Options] 

Note that only the token file is a required argument. Other options are explained below:
    
    Arguments (Required):
    
        TOKEN_FILE              Specify the filename for the text file that contains your Asana Personal 
                                Access Token as "filename.txt"
    
    Options:
        
        -w, --write     TEXT    Specify a filename for the output spreadsheet as "Filename.xlsx".
                                By default, the output spreadsheet is 
                                "DATE-TIME-Implementation_Dashboard.xlsx", 
                                where DATE-TIME is the current date and time.
        
        -t, --template  TEXT    Specify the name of the Template Task List to retrieve from Asana as 
                                "[TEMPLATE] Template Name". The default template is "[TEMPLATE] Outpatient 
                                Services".
        
        -s, --strict            Use strict matching against template task names instead of matching by 
                                prefixed task numbers [xx].
        
        --help                  Show this message and exit.

### Running in Strict Mode

Strict Mode allows strict matching for task names - this means that task names will be matched against the template tasks using the entirety of the name, and must be an *exact* match. For example, the following would be a match:

- Template Task Name: “Hand-off from Implementation to Account Management”
- Project Task Name: “Hand-off from Implementation to Account Management”

The default matching (not in Strict Mode) is by the task prefix as specified in between brackets at the beginning of the task name. For example, the following would be a match:

- Template Task Name: “[5A] Hand-off from Implementation" 
- Project Task Name: “[5A] Hand-off to Account Management”

You can use older templates (that do not use the task prefixing) to generate the Implementation Dashboard spreadsheet, by specifying the template name using the `-t` option, and specifying Strict Mode via the `-s` option. For example:

    $ idash token.txt -t "[TEMPLATE] Implementation Task List" -s

### Asana Template Usage
#### Metadata Section
The following "tasks" in the Metadata section should be changed accordingly:
* Implementation Start Date task example:
  * [0A] 09/10/2016
* Go-Live Target Date task example:
  * [0B] 12/10/2016
* Template specification (for different Templates) example:
  * [0C] Outpatient Services

#### Actual Completed Start Date subtask
For every task that is completed you may create a "Actual Completed Start Date" subtask when the checked-off date in Asana does not represent the true completion date of the task. 
* Example: Actual Completed Date: xx/xx/xxxx

