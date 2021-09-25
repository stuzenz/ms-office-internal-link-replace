# Office files bulk update tool

This tool has been built out to support a low effort migration from one document management system to another.
It was developed to cover programme needs when the programme decided the knowledge base would be moved to a new location to allow 'one team' where both internal staff and vendors could easily work with the same files. This was not easily possible with the client internal sharepoint-based knowledge management system or with the prime vendor sharepoint site

This tool has been developed in python with a behaviour driven design approach

"""
Gherkin step implementations for presentation-level features
"""

## given =================================
@given('Starting the bulk update tool')


## when =================================
@when('Choosing the option to start with a new clean xlsx template is offered')

@when('Choosing the option to open a previously generated excel template

@when('the user wants to collect all internal links for review


## then =================================
@then('An empty template is created and saved to the working directory') 

@then('the user is directed a dialog to open the already populated xlsx template, and the file is opened for editing')

@then('the base directory is scanned and a list of documents, locations, and the internal links in the documents is shown in the open excel sheet'