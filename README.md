# iDash

Generate the Implementation Dashboard from projects and tasks in Asana


# Installation

Simply run:

    $ pip install .


# Usage

To use it:

    $ idash [OPTIONS] TOKEN_FILE
    
    Arguments:
    
        TOKEN_FILE      TEXT    Specify a filename for your Asana Personal Access Token as "token.txt"
    
    Options:
        
        -w, --write     TEXT    Specify a filename for the output spreadsheet as "Filename.xlsx".
                                By default, the output spreadsheet is "Implementation_Dashboard.xlsx".
        
        -t, --template  TEXT    Specify the name of the Template Task List to retrieve from Asana as 
                                "[TEMPLATE] Name". The default template is "[TEMPLATE] Outpatient Services".
        
        -s, --strict            Use strict matching against template task names instead of matching by 
                                prefixed task numbers [xx].
        
        --help                  Show this message and exit.

