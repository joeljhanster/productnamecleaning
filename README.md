# productnamecleaning
HiMart manages more than 6000 products. In order to facilitate better data analysis, it is vital that all the product names come in a standard format (Brand, Product Name, Volume, Quantity). This repository is a data cleaning script that I created to make the boring manual work a much faster process. 

### Thought Process ###
After attempting to manually change the product names one by one. The tedious process was far more dreadful than I imagined. As such, I decided to write a Python Script to break down the original product names into 4 parts, individually extracting the different components from the original names and placing them into specified cells in the new excel sheet. Each time an information is extracted, it is cut out from the original name (thus the new processing name is shorter). Finally, I use the CONCATENATE function in Excel to join all these components together, and used TRIM to remove all additional spaces.

### Extracting Volume ###
This is the easiest to extract because volumes have fixed prefixes, such as ml, kg etc. Therefore, using RegEx, it was an easy process to extract the volume.

### Extracting Brand ###
This process is slightly more tedious. But after all the manual work, there is a pattern I saw in which a certain brand has only that many abbreviations. More often than not, the namings by suppliers provide similar abbreviations (e.g. RIB => Ribena). Using a dictionary stored in the `brand_list.txt` file, I constantly update this file whenever I match a new abbreviation to the actual brand name.

### Extracting Quantity ###
Assuming that quantity involves `int` values, I will look through the remainings of the product names to see if there are any `int` values in there. With that, I will do the same as what I did for extracting brand, in which I have a dictionary to store all the potential abbreviations for quantities, e.g. "6's" => "(6s)".

### Extracting Product Name ###
After extracting all of the above, it becomes much easier to just head over to Excel to copy and paste the remainings of the product name into the respective product name column. In fact, this last extraction does not require the use of Python.

### Final Product ###
<img src="https://github.com/joeljhanster/productnamecleaning/blob/master/cleaning.png" />

NOTE: Due to confidentiality, I will not upload the excel sheet that required cleaning.
