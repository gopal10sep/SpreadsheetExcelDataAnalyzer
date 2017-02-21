import xlrd
import numpy as np
import matplotlib.pyplot as plt

# Open the workbook
xl_workbook = xlrd.open_workbook("sample.xlsx") 


# Grabbing the Required Sheet by index 
xl_sheet = xl_workbook.sheet_by_index(1)

#Initialising the Lists Required
list_application_year = []
list_publication_year = []
complete_list=[]
years =[]
app_year_freq_list=[]
pub_year_freq_list=[]

num_cols = xl_sheet.ncols   #Number of columns
for row_idx in range(2, xl_sheet.nrows):    # Iterate through rows

    for col_idx in range(11, 12):  # Iterate through Application Year column
        cell_obj = xl_sheet.cell(row_idx, col_idx)  # Get cell object by row, col
        application_year = str(cell_obj)[7:11]
        list_application_year.append(application_year)
        complete_list.append(application_year)
        
    for col_idx in range(12, 13):  # Iterate through Publication Year column
        cell_obj = xl_sheet.cell(row_idx, col_idx)  # Get cell object by row, col
        publication_year = str(cell_obj)[7:11]
        list_publication_year.append(publication_year)
        complete_list.append(publication_year)

#Making Dictionaries for the Lists obtained.
dict_app_year = {x:list_application_year.count(x) for x in list_application_year}
dict_pub_year = {x:list_publication_year.count(x) for x in list_publication_year}

max_year = int(max(complete_list))
min_year = int(min(complete_list))
while min_year<= max_year:
	#Constructing the Years List
	years.append(min_year)
	#Constructing Application Year Freq List
	if dict_app_year.get(str(min_year)) != None:
		app_year_freq_list.append(int(dict_app_year.get(str(min_year))))
	else:
		app_year_freq_list.append(0)
	#Constructing Publication Year Freq List
	if dict_pub_year.get(str(min_year)) != None:
		pub_year_freq_list.append(int(dict_pub_year.get(str(min_year))))
	else:
		pub_year_freq_list.append(0)

	min_year = min_year +1 

#Plotting the Graph
N = len(years)
ind = np.arange(N)  	#The x locations for the groups
width = 0.35       		#The width of the bars
fig, ax = plt.subplots()
rects1 = ax.bar(ind, app_year_freq_list, width, color='r')
rects2 = ax.bar(ind + width, pub_year_freq_list, width, color='y')

# add some text for labels, title and axes ticks
ax.set_ylabel('Frequency')
ax.set_title('Number of Applications And Publications in an Year')
ax.set_xticks(ind + width / 2)
ax.set_xticklabels(years, rotation='vertical')
ax.legend((rects1[0], rects2[0]), ('No. of Applications', 'No. of Publications'))
#Attach a text label above each bar displaying its height
def autolabel(rects):    
    for rect in rects:
        height = rect.get_height()
        if height != 0 :
        	ax.text(rect.get_x() + rect.get_width()/2., 1.05*height,'%d' % int(height),ha='center', va='bottom')
autolabel(rects1)
autolabel(rects2)

# Tweak spacing to prevent clipping of tick-labels
plt.subplots_adjust(bottom=0.15)
axes = plt.gca()
axes.set_ylim([0,20])
plt.show()
