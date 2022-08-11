# Pseudocode

* (1) Create a database with the various recipes that I want to choose from
    * Tag different recipes based on style (i.e. Mexican, Pasta, etc.)
    * Link ingredients to those recipes

    ## CODE
    - Put this in a spreadsheet 
    - Can export to pandasDF
    - And DF.tohtml() to export to a website for Maria to see


* (2) Create a user-input model for selecting weekly recipes
    * Add a filter to manually select some recipes
    * Add a random select button if not wanting to choose

    ## CODE
    - Use MsgBox to allow user to input a specific recipe id/number of recipes to generate
    - Dim empty variables for each item that needs to be grabbed
    - Randomly choose a recipe_id
    - Save the values of that recipe_id and print in the destination worksheet

    - Could also use pandas '.sample(5)' to get a full random 5
    - Could also use pandas '.sample(1)' to get a random 1


* (3) Create a grocery-list database
    * Gather the ingredients needed for the recipes
    * Prompt user for how many ingredients currently have**
    * Adjust required ingredients list to account for the ones that are already there

    ## CODE
    - Set up different sections for meat, bread, produce, etc
    - Store ingredient columns as decimal-comptabile number types
    - Find sums for each ingredient column
    - Use similar code as the original macro for-loop to gather string information
    - Display the information in the proper section

    
* **OPTIONAL** (4) Create a display website to display actual recipes instructions
    * Either type up recipes to be referenced by a website
    * Or take pictures of printed recipes from cookbooks (copyright infringement?)
    * Then upload onto the website