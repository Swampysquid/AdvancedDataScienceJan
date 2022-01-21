
# **Excel Explainer: VLOOKUP**
**By: Eythan Jenkins**

## What is VLOOKUP
Excel's VLOOKUP tool is a search function for "when you need to find things in a table or a range by row." (support.microsoft.com) Essentially, you give VLOOKUP input, Excel searches for that input in the chart, and finally it will give you the value associated with that input. I like to think of it as being similar to a hash table or dictionary: you provide the VLOOKUP function with your 'key', and it returns the 'value' of that key.

## Why to use VLOOKUP
A search function like VLOOKUP makes it easy to find values fast within an Excel sheet. Instead of having to maually navigate the sheet for your data of interest, you can simply search up the term you are looking for, and the value will be returned to you. An example of this can be seen below:
![Image](VLOOKUPpage1.JPG)
In the above example, the user is prompted to put a 'Major' into the blue cell E3. After doing this, the orange and green fields in cells E5 and E7 will return values extracted from the VLOOKUP function. While this example looked at the functionality on a small scale, VLOOKUP would be extremely useful in scenarios where there are hundreds or thousands of data points to look through. For instance, an excel sheet for a supermarket may contain all of the products, the quantity in store, and the retail price. Instead of having to go through a list of all the products, and manually find one's price, VLOOKUP can be used to expedite the process: all you need is the name of the item. 

## How to Use VLOOKUP
A sample VLOOKUP function will likely look something like this:

![Image](VLOOKUPFunction1.JPG)

To explain how to use the VLOOKUP function, I am going to split it into 4 parts:

### What You are Looking For/ Key

![Image](VLOOKUPFunction2.JPG)

### Range to Search

![Image](VLOOKUPFunction3.JPG)

### Which Column to Return From

![Image](VLOOKUPFunction4.JPG)

### Approximate or Exact Key Match

![Image](VLOOKUPFunction5.JPG)

## Data Used

## References
#### VLOOKUP Function Basics:
[support.microsoft.com](https://support.microsoft.com/en-us/office/vlookup-function-0bbc8083-26fe-4963-8ab8-93a18ad188a1)
#### VLOOKUP Approximate versus Exact Values:
[extendoffice.com](https://www.extendoffice.com/documents/excel/2443-excel-vlookup-exact-approximate-match.html)


- Bulleted
- List

1. Numbered
2. List

**Bold** and _Italic_ and `Code` text

[Link](url) and ![Image](src)
