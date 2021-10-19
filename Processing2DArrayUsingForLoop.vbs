Option Explicit

Dim PhoneBook(2, 4) ' 3 cols 5 rows
Dim colLowerIndex, colHigherIndex, rowLowerIndex, rowHigherIndex
Dim colIndex, rowIndex, searchName, matchedrow, ri, ci, found

PhoneBook(0,0) = "Tilak"
PhoneBook(1,0) = "Boston"
PhoneBook(2,0) = "111-111-0000"

PhoneBook(0,1) = "Swetha"
PhoneBook(1,1) = "Michigan"
PhoneBook(2,1) = "111-111-0001"

PhoneBook(0,2) = "Megha"
PhoneBook(1,2) = "Chicago"
PhoneBook(2,2) = "111-111-0002"

PhoneBook(0,3) = "Medha"
PhoneBook(1,3) = "Newyork"
PhoneBook(2,3) = "111-111-0003"

PhoneBook(0,4) = "Peter"
PhoneBook(1,4) = "Seattle"
PhoneBook(2,4) = "111-111-0004"

colLowerIndex = 0
colHigherIndex = UBound(PhoneBook, 1) ' 1st dimension of array
rowLowerIndex = 0
rowHigherIndex = UBound(PhoneBook, 2) ' 2nd dimension of array

'using nested for loop
for ri = rowLowerIndex to rowHigherIndex
    for ci = colLowerIndex to colHigherIndex
        WScript.Echo PhoneBook(ci, ri)
    Next
    WScript.Echo "-----"
Next
