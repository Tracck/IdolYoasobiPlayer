# IdolYoasobiPlayer
The files and code that I used to render "Idol" on Excel

If the excel file doesn't contain the marco then paste the code from "ExcelVBAMarcoCode.txt" in
Alt + F11 to open the marco coding window in excel

The python script isn't completely made by me, I had to go to stackoverflow and copy some of them 
to save my sanity. 

There will be a section in the python script which may look like it is made from YandereDev, it is 
because I can't pass an array into excel vba macro function, therefore I had to custom made 15
different choices for 15 different parameters (which is very limited, the code might break if you
go for a resolution > 144x256, however, I tried 360x640 and it is still fine)

Lastly, the Image Loading algorithm isn't the best, feel free to improve it.

I won't ask you to credit me for the code but it would make my day if you do!

<b>CERTAIN LIMITATIONS YOU MIGHT WANT TO KNOW:<b\>
- Excel can only handle 64000 different cell formats at once, if you exceed that limit
the code breaks. You need to clear all cells' formats AND RESTART EXCEL

- The Marco function/procedure can only hold as many as 60 parameters, if you do the
functions/procedures my way then having 57 different parameters is your limit (3
parameters R, G, B for color)

- The maximum cells a Range object in Excel can hold is about 4000 cells. That's why
I added the limit of only having 20 different CellRanges (a CellRange is like "A1:C3"
nr refers to NameRange btw) per Range object. However it is a lazy workaround as
sometimes 3 CellRanges can be over 4000 cells and sometimes 50 CellRanges can be as 
little as 20 cells. You might want to make a function to calculate how many cells
a CellRange is holding.

- If you read the code you will realize that I screenshot my own screen to capture
a frame, remember to turn off any social media apps (or at least the notifications)
before you go afk

- There is no focus function to focus back on Excel in my code when it restart so
be careful before you go afk.

- Remember to name your excel file extension as .xlsm so it saves the macro code
