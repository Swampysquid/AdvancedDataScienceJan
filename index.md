
# **Excel Explainer: VLOOKUP**
**By: Eythan Jenkins**

## What is VLOOKUP
Excel's VLOOKUP tool is a search function for "when you need to find things in a table or a range by row." (support.microsoft.com) Essentially, you give VLOOKUP input, Excel searches for that input in the chart, and finally Excel gives you the value associated with that input. I like to think of it as being similar to a hash table or dictionary: you provide the VLOOKUP function with your 'key', and it returns the 'value' of that key.

## Why to use VLOOKUP
A search function like VLOOKUP makes it easy to find values fast within an Excel sheet. Instead of having to manually navigate the sheet for your data of interest, you can simply search up the term you are looking for, and the value will be returned to you. An example of this can be seen below:
![Image](VLOOKUPpage1.JPG)
In the above example, the user is prompted to put a 'Major' into the blue cell E3. After doing this, the orange and green fields in cells E5 and E7 will return values extracted from their own VLOOKUP functions. While this example looked at the functionality on a small scale, VLOOKUP would be extremely useful in scenarios where there are hundreds or thousands of data points to look through. For instance, an excel sheet for a supermarket may contain all of the products, the quantity in store, and the retail price. Instead of having to go through a list of all the products, and manually find one's price, VLOOKUP can be used to expedite the process: all you need is the name of the item. 

## How to Use VLOOKUP
A completed VLOOKUP function will look something like this:

![Image](VLOOKUPFunction1.JPG)

To use, the VLOOKUP function is written into a cell. After writing your initial "=VLOOKUP()", there will be four variables of interest to enter into the function parentheses:

### What You are Looking For/ Key

![Image](VLOOKUPFunction2.JPG)

This first value refers to what you are searching for / 'the key.' This value can be hardcoded in (I could have put 'Economics' in place of 'E3'). However, this makes it a little less flexible, since you will always have to go back into the function to make a change. Instead, by using a cell, I can just go into that cell and change that value.

![Image](VLOOKUPpage1_2.jpg)

### Range to Search

![Image](VLOOKUPFunction3.JPG)

This second zone is the scope of the VLOOKUP function. This is where the VLOOKUP function is going to look for the key and the corresponding values. In this example, the keys and their corresponding values exist from A2:C14, so I enter that into the function for the range. The VLOOKUP function looks for the key in the first column of the range, so the function will be searching through the A column to find the matching value.

![Image](VLOOKUPpage1_3.jpg)

### Which Column to Return From

![Image](VLOOKUPFunction4.JPG)

The third part of the VLOOKUP function refers to which column of the range should be returned to the function as the found value. So if I want the value from the B column, I enter 2, since that is the second column within the range of my function. By doing this, it returns to me the proper letter grade. For the numeric value, I would change 2 to 3, since the numeric values are stored in the third column of the range.

![Image](VLOOKUPpage1_4.jpg)

### Approximate or Exact Key Match

![Image](VLOOKUPFunction5.JPG)

The final part of the VLOOKUP function is a boolean value: 'False' determines whether the key needs to perfectly match, and 'True' permits there to be an 'approximate' match. By 'approximate,' VLOOKUP "returns the next largest value that is less than your specific lookup value." (extendoffice.com) Because of how it is defined, it requires the dataset to be sorted in an ascending manner as it goes down the column. Because the value is False in this case, the search key must match exactly, else it will return a N/A value to the return fields.

A time when it may be helpful to set it to true is if you have a dataset with multiple thresholds. Instead of needing a value for every possible dataset, you can use the approximate value and VLOOKUP will be able to return a value despite it not being an exact match. Another example demonstrates the usefulness of this variable:

![Image](VLOOKUPpage2.JPG)

In this example, my VLOOKUP functions are set to 'True' instead of 'False'. My 'key' value of 3999, is not in the dataset that VLOOKUP is looking through. Despite this, it is still able to return values to the function. These values are derived from row 11, which stored a value of 3000 in its first column. This is because 3000 is the next largest value, that is less than the specific lookup value. So even though 4000 is closer to 3999 than 3000 is, Excel will use the values in the row containing 3000.

![Image](VLOOKUPpage2_2.JPG)

## Data Used

[Webpage with Excel File for Download](https://github.com/Swampysquid/AdvancedDataScienceJan/blob/gh-pages/VLOOKUPExplainer.xlsx)

## References
#### VLOOKUP Function Basics:
["VLOOKUP Function" Retrieved from: support.microsoft.com on (1/20/22)](https://support.microsoft.com/en-us/office/vlookup-function-0bbc8083-26fe-4963-8ab8-93a18ad188a1)
#### VLOOKUP Approximate versus Exact Values:
["How To Use Vlookup Exact And Approximate Match In Excel?" Retrieved from: extendoffice.com on (1/20/22)](https://www.extendoffice.com/documents/excel/2443-excel-vlookup-exact-approximate-match.html)



