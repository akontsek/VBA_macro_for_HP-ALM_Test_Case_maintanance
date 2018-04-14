1.) REQUIREMENTS:
Here are the mandatory columns needed for the ALM Upload Add-in for Excel (these are only the basic, you may add more):

TEST TYPE	
SUBJECT	
DESIGNER	
OWNERSHIP	
TEST LEVEL	
TEST NAME	
STEP NAME (DESIGN STEPS)	
DESCRIPTION (DESING STEPS)	
EXPECTED (DESING STEPS)

HP-ALM Upload AddIn for Excel
https://www.youtube.com/watch?v=Lz9JEX_tr-k - tutorial


2.) WHO NEEDS THIS?
	
This VBA macro is useful for those testers who need to maintain the existing test cases and prepare them for their next use. 
If there are just little changes needed in the test basis, but multiplied many times, then it's easier to handle them in Excel first and then upload them to HP-ALM via the Upload Add-in.

Examples of use cases:
-if a developed function (that's been already tested) is about to implement in several domains (eg.: markets, target-groups, user-profiles) -> the Test Cases require market-specific inputs


3.) WHAT TO CHANGE?
a.) The path in HP-ALM (column 'Subject'):    
eg.: 
"Main folder\Sub-folder\TestCase_01" 
"Main folder\Sub-folder\TestCase_02"
"Main folder\Sub-folder\TestCase_03"

b.) Parts or the whole Test Name:
eg.:
"TestName_01"
"TestName_02"
"TestName_03"

c.) Test inputs, Test data, or any such kind of a variable:


4.) SETTINGS FOR THE MACRO

According to your initial test sample you need to set the parameters that are going to needed for the multiplication.

Basically what the macro does is:
1.) getting the new values from manually set arrays (described below)
2.) copy&paste the sample Test Case(s) 
3.) do the first replacements in the text
4.) do the second...to...last replacements in the text 
5.) repeats step 2.-4.  as many times as it was set in the arrays 
6.) put the header for the first row (needed for the AddIn)


5.) PARAMETERS:
From the original sample Test Cases you have to define the variables that are going to be changed, and add the new values to each array in the VBA format.

eg.:
"\TestCase_01"    -->  arraypath = Array("\TestCase_02", "\TestCase_03","\TestCase_you-name-it")
"TestName_01"    -->  arrayname = Array("TestName_new_A", "TestName_new_B", "TestName_new_C")
"aaa"    -->  arrayinput = Array("bbb", "ccc", "ddd")
"111"    -->  arraydata = Array("222", "333", "444")

