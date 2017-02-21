# SpreadsheetExcelDataAnalyzer

Working:<br />
1.Using xlrd package we are importing the excel sheet and reading its contents.<br />
2.Iterating through Application Year column, I am appending the years extracted in to a list list_application_year[].<br />
3.Similarly,Iterating through Publication Year column, I am appending the years extracted in to a list list_publication_year[].<br />
4.I also made a list having the complete data complete_list[].<br />
5.I transform the lists into dictionaries (dict_app_year, dict_pub_year), beacause dictionaries are of the form key:value.<br />
6.While ietrating through the min_year to max_year in the complete list, I make two freq lists (app_year_freq_list[],pub_year_freq_list[]) corresponding to the year and their applications and publications using the dicitionaries earlier created.<br />
7.Rest is simple graph plotting as readable by the comments included in the code. <br />

Read Me:<br />
1. Extract the folder contents.<br />
2. The program reads the given excel file “sample.xlsx” and gives a relation between no. Of applications and publications in a given year.<br />
3. Modules and Python Packages required:<br />
• sudo pip install -U pip setuptools<br />
• sudo pip install xlrd<br />
• sudo apt-get install python3-tk<br />
• sudo pip install matplotlib<br />
• sudo pip install numpy<br />
4.Run the program: python excel.py<br />
