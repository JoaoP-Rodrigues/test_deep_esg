# test_deep_esg
Repository created for a skills test.

In order to solve this challenge, I wrote the code in several steps, and in each one of them I was printing the values to know if I was handling them correctly.
First, I created the reading part of the files received via input, and tried some simple slicing and type conversions.
Then I created the output file and recorded some random cells to test if everything was right.
After that, I made the code logic:
Open the first file (chart_of_accounts), get a first value from "A" column and assign in a variable.
So, use this variable to search in second file (general_ledgers) for a equal value using a "for loop". 
Finding, get a respective value from "B" Column and increment in a temporary variable.
Next, the value of this variable will be write in the output file variable, in the respective row.
So, the code return to the start of loop.

In the next step, all these values, assigned on the variable that open the file, will be write and save in the "xlsx" file.

In the last step, there will be the combination of branches of tree.
For this, the code get a first cell with value from B column equal zero and assign to the new variable.
This variable will be a search key for second loop.
So, the code use this variable for search for all branches with value equal the key, with a length limitation to get a control, and not get a different branch.
If the value are equals, it will saved in a accum variable, and write in the output file.
Lastly, the output file is saved and closed.

----------------------------------------------´

All the development was made in Windows 10 environment.
I not know so much from Linux to test this code  in your environment.

All development was done in Windows 10 environment.
I don't know much about Linux to test this code in your environment.
I did some tests, using files bigger than the ones provided, and everything worked perfectly. 

----------------------------------------------
Notes:
    ->For this code to work perfectly, it needs a directory called 'input', with two saved files:
        ---->chart_of_accounts.xlsx
        ---->general_ledger.xlsx
        ->If the files do not exist, an error message will be displayed by the code.
    ->These files also need to be completed. If not, the code will work, but the output file will be blank.
    ->You must have write permission on the directory to save the output file. If there is no such permission, an error message will be displayed by the code. 
