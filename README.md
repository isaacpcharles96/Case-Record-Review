The purpose of this code was to generate a table of randomly assigned cases for review by staff each month. The process was originally performed in Excel using a series of formulas that were rather time consuming. 
Our process for randomization was also deeply flawed, involving using the RAND() function and sorting it each time for a truly random selection. 
Python excels in performing repetitive and randomized tasks like this and therefore I decided to move this process to Python.
Using libraries such as pyodbc for SQL connectivity, pandas for dataframe manipulation, and timedelta, I was able to recreate this reporting process using Python instead.
