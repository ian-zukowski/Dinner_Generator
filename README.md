# Dinner_Generator


## DESCRIPTION OF MACROS
* Recipe_Generator
* Random_Recipe
* Grocery_List

## FUTURE IMPROVEMENTS
* Allow user to input a desired recipe_type in Recipe_Generator
    <!-- 'If not requesting a specific recipe type 0 first
    'If you want random, type "Random"
    'If you want specific, type "___"
        'Will want to display list of distinct food_type options -->
* Consolidate empty boxes in Grocery_List
    <!-- 'Check if A2 is empty -- if so set B2 to A2, C2 to B2, ... -->
* Account for what is already in the fridge in Grocery_List
    <!-- 'Run through value fields and have separate InputBox
        '"Do you have ____?" - for strings
        '"Do you have X _______?" for integers 
        ' For duplicates, check if B2=A2, and change InputBox to "Do you have another ___?"-->
* Add a "I have 'x' ingredient... give me a recipe using that" macro
    <!-- 'Will be super similar to the desired recipe_type improvement
    'Would have to try to account for uppercase/lowercase for this -->
* Website upload so family members can easily view
    <!-- 'Will need to see if there is a to_html type function for excel
    'Otherwise will need to export to a pandas df
    'Then to_html from the df -->

## CODE BEHIND THE MACROS
For more in depth pseudocode, be sure to check out the "Pseudocode" folder in the above resources, or click the relevant link provided.

* [RecipeGenerator](Pseudocode/macro_pseudocode.txt)
    - Use MsgBox to allow user to input X amount of recipes to generate
    - Dim empty variables for each item that needs to be grabbed
    - Start a for loop with X iterations in the database
    - Randomly choose a recipe_id
    - Save the values of that recipe_id and print in the destination worksheet

* [RandomRecipe](Pseudocode/single_recipe_macro_pseudocode.txt)
    - Use MsgBox to allow user to input a specific recipe_id to import (or allow a random reroll)
    - Dim empty variables for each item that needs to be grabbed
    - Locate the desired recipe_id (or pick a random recipe_id)
    - Save the values from that recipe_id and print in the destination worksheet

* [GroceryList](Pseudocode/grocery_macro_pseudocode.txt)
    - Set up different sections for meat, bread, produce, etc
    - Store ingredient columns as decimal-comptabile number types
    - Find sums for each ingredient column
    - Use similar code as the original macro for-loop to gather string information
    - Display the information in the proper section