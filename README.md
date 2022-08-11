# Dinner_Generator
Hello, and welcome to the greatest development in the culinary software field since sliced dataframes! In this Github repo you can find the .xlsm file which I've been using since June 2022 to help me stay consistent in cooking homemade meals instead of always eating out. I was inspired to create this project when I realized I had been ordering takeout much more frequently during the pandemic, and then even moreso during my wife's first pregnancy when she would get cravings for fast food. 

I've enjoyed cooking ever since I graduated college, and I understood that my decline in home cooking stemmed mainly from not having a coherent plan for the week ahead when I went grocery shopping. This would then create a cycle of:

1. Not being able to think of more than 7 or 8 recipes I could do with the ingredients in my pantry
2. Doing the same set of recipes over and over
3. Using Kroger Clicklist to order groceries, and seeing all my 'Previously Bought' items first and thinking "You're right Clicklist, we DID use zucchini every week this past month, so we probably need to restock those"
4. Loop back to 1, where I have 3 zucchini that I have to use up and can only think of a couple recipes that I've done with them recently

Occassionally I'd find a recipe I wanted to try out, but that always required a trip up to the grocery store to grab several ingredients which we didn't currently have. So with that cycle in mind, I figured I really ought to make a list of all the different recipes I could choose from before going grocery shopping. That way I would have the ingredients I needed, without having extras that force me into doing the same type of meals over and over. 

And here we reach the purpose of this repo: I've created a macro-enabled file which houses 30+ dinner ideas for me to generate before going grocery shopping. And after generating the list of meals I'd like to cook, it also has the ability to generate a grocery list which I can easily print to one sheet of paper in landscape. This has made my cooking process MUCH more streamlined and diverse, which also has the side benefit of giving me more time with my new son!

## DESCRIPTION OF MACROS
* Recipe_Generator
    - This macro is located in the "Recipe_Generator" worksheet of the file. It will prompt a user to input the necessary amount of recipes for their desired time period, then randomly select that amount of recipes from the My_Recipes worksheet to display as a separate list.
* Random_Recipe
    - This macro is located in the "Recipe_Generator" worksheet of the file. It will first prompt a user to input the recipe to be replaced. Then it will prompt the user to input a desired or random recipe from the My_Recipes worksheet. The results of this second prompt will overwrite the singular original recipe that was chosen before.
* Grocery_List
    - This macro is located in the "Grocery_List" worksheet of the file. It will access the generated list of recipes located in the "Recipe_Generator" worksheet, and then print the results in separated sections (roughly based on locations in Kroger aisles)

## FUTURE IMPROVEMENTS
* Add the actual recipes I use into the database so others can see how to make the meals (or I can double check the source recipe)
    <!-- No code required -- just a ton of data input -->
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
* Add a "I'm feeling adventerous" macro to access new recipes that aren't in the tried-and-true database
    <!-- 'Will need a disclaimer to do AFTER generating other recipes for the week
    'Won't need to use any ingredient variables for this one
    'Do something similar to row count to find the last row, which will contain this new recipe
    'Probably want to separate it even further -- last_row holds titles: "Recipe Name", "Recipe Source", and last_row+1 holds the values -->
* Website upload so my wife can easily view/veto the schedule
    <!-- 'Will need to see if there is a to_html type function for excel
    'Otherwise will need to export to a pandas df
    'Then to_html from the df -->

### CODE BEHIND THE MACROS
For pseudocode, be sure to check out the "Pseudocode" folder. Clicking the link will lead to the actual code used for the macros.

* [RecipeGenerator](Actual_Code/recipe_gen.txt)
    - Use MsgBox to allow user to input X amount of recipes to generate
    - Dim empty variables for each item that needs to be grabbed
    - Start a for loop with X iterations in the database
    - Randomly choose a recipe_id
    - Save the values of that recipe_id and print in the destination worksheet

* [RandomRecipe](Actual_Code/single_recipe.txt)
    - Use MsgBox to allow user to input a specific recipe_id to import (or allow a random reroll)
    - Dim empty variables for each item that needs to be grabbed
    - Locate the desired recipe_id (or pick a random recipe_id)
    - Save the values from that recipe_id and print in the destination worksheet

* [GroceryList](Actual_Code/grocery_list.txt)
    - Set up different sections for meat, bread, produce, etc
    - Store ingredient columns as decimal-comptabile number types
    - Find sums for each ingredient column
    - Use similar code as the original macro for-loop to gather string information
    - Display the information in the proper section