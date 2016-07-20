# iDash

A command line tool that generates the Implementation Dashboard spreadsheet from projects and tasks in Asana.


# Installation

Check to see if you have Pip installed on your system. You can test this by running:

    $ pip -v

If the command is not recognized, you do not have Pip. Pip generally comes bundled with a Python install. To install Python, including Pip and other setup tools: https://www.python.org/downloads/ 

Once pip has been installed, clone this repository to your desktop, then navigate to the root directory. 

Simply run:

    $ pip install .
    
This command will find and install all the Python packages needed for the script to run. 

# Usage

To use it:

    $ idash TOKEN_FILE [Options] 
    
    Arguments (Required):
    
        TOKEN_FILE      TEXT    Specify a filename for your Asana Personal Access Token as "filename.txt"
    
    Options:
        
        -w, --write     TEXT    Specify a filename for the output spreadsheet as "Filename.xlsx".
                                By default, the output spreadsheet is "Implementation_Dashboard.xlsx".
        
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
