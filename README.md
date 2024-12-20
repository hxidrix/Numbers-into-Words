 # Custom Formula for Converting Numbers into Words in Excel

This project involves creating a custom formula in Excel that converts numeric values into written words, such as "1000" being converted into "One Thousand". The formula is developed using VBA (Visual Basic for Applications) code, and it simplifies the process of displaying numbers as words within Excel. Here's how you can implement and use the custom formula:

**Method:**

1. **Copy the VBA Code**: 
   - First, copy the VBA code provided for converting numbers into words.
```
Option Explicit

'Main Function

Function SpellNumber(ByVal MyNumber)

Dim Dollars, Cents, Temp

Dim DecimalPlace, Count

ReDim Place(9) As String

Place(2) = " Thousand "

Place(3) = " Million "

Place(4) = " Billion "

Place(5) = " Trillion "

' String representation of amount.

MyNumber = Trim(Str(MyNumber))

' Position of decimal place 0 if none.

DecimalPlace = InStr(MyNumber, ".")

' Convert cents and set MyNumber to dollar amount.

If DecimalPlace > 0 Then

Cents = GetTens(Left(Mid(MyNumber, DecimalPlace + 1) & "00", 2))

MyNumber = Trim(Left(MyNumber, DecimalPlace - 1))

End If

Count = 1

Do While MyNumber <> ""

Temp = GetHundreds(Right(MyNumber, 3))

If Temp <> "" Then Dollars = Temp & Place(Count) & Dollars

If Len(MyNumber) > 3 Then

MyNumber = Left(MyNumber, Len(MyNumber) - 3)

Else

MyNumber = ""

End If

Count = Count + 1

Loop

Select Case Dollars

Case ""

Dollars = "No Dollars"

Case "One"

Dollars = "One Dollar"

Case Else

Dollars = Dollars & " Dollars"

End Select

Select Case Cents

Case ""

Cents = " and No Cents"

Case "One"

Cents = " and One Cent"

Case Else

Cents = " and " & Cents & " Cents"

End Select

SpellNumber = Dollars & Cents

End Function


' Converts a number from 100-999 into text

Function GetHundreds(ByVal MyNumber)

Dim Result As String

If Val(MyNumber) = 0 Then Exit Function

MyNumber = Right("000" & MyNumber, 3)

' Convert the hundreds place.

If Mid(MyNumber, 1, 1) <> "0" Then

Result = GetDigit(Mid(MyNumber, 1, 1)) & " Hundred "

End If

' Convert the tens and ones place.

If Mid(MyNumber, 2, 1) <> "0" Then

Result = Result & GetTens(Mid(MyNumber, 2))

Else

Result = Result & GetDigit(Mid(MyNumber, 3))

End If

GetHundreds = Result

End Function


' Converts a number from 10 to 99 into text.


Function GetTens(TensText)

Dim Result As String

Result = "" ' Null out the temporary function value.

If Val(Left(TensText, 1)) = 1 Then ' If value between 10-19...

Select Case Val(TensText)

Case 10: Result = "Ten"

Case 11: Result = "Eleven"

Case 12: Result = "Twelve"

Case 13: Result = "Thirteen"

Case 14: Result = "Fourteen"

Case 15: Result = "Fifteen"

Case 16: Result = "Sixteen"

Case 17: Result = "Seventeen"

Case 18: Result = "Eighteen"

Case 19: Result = "Nineteen"

Case Else

End Select

Else ' If value between 20-99...

Select Case Val(Left(TensText, 1))

Case 2: Result = "Twenty "

Case 3: Result = "Thirty "

Case 4: Result = "Forty "

Case 5: Result = "Fifty "

Case 6: Result = "Sixty "

Case 7: Result = "Seventy "

Case 8: Result = "Eighty "

Case 9: Result = "Ninety "

Case Else

End Select

Result = Result & GetDigit(Right(TensText, 1)) ' Retrieve ones place.

End If

GetTens = Result

End Function


' Converts a number from 1 to 9 into text.

Function GetDigit(Digit)

Select Case Val(Digit)

Case 1: GetDigit = "One"

Case 2: GetDigit = "Two"

Case 3: GetDigit = "Three"

Case 4: GetDigit = "Four"

Case 5: GetDigit = "Five"

Case 6: GetDigit = "Six"

Case 7: GetDigit = "Seven"

Case 8: GetDigit = "Eight"

Case 9: GetDigit = "Nine"

Case Else: GetDigit = ""

End Select

End Function
```

2. **Open the Excel Workbook**: 
   - Open the workbook where you want to use the custom formula.

3. **Access the VBA Editor**:
   - Right-click on the sheet tab where you want to use the formula.
   - Select **View Code** to open the VBA editor.

4. **Insert a Module**:
   - In the VBA editor, click on the **Insert** tab, then select **Module** to insert a new module.

5. **Paste the Code**:
   - Paste the copied VBA code into the module window.

6. **Declare the Formula**:
   - Ensure the formula is declared properly, in this case, using the `SpellNumber` declaration. This makes it available for use in your Excel workbook.

7. **Close the VBA Editor**:
   - After pasting the code and ensuring the formula is declared, close the VBA editor.

8. **Use the Formula**:
   - Now, in any cell of the sheet, you can use the custom formula to convert numbers to words.
   - The formula should be written as:
     ```
     =SpellNumber(A1)
     ```
     Replace **A1** with the reference of the cell containing the number you want to convert.

9. **Currency Formatting**:
   - The formula automatically converts the number into words, and it will also work for converting numbers into dollar currency, such as "One Thousand Dollars" for "1000".



**Changing the Currency:**

If you wish to change the currency from dollars to another type, follow these steps:

1. **Open the VBA Editor**:
   - Press **Alt + F11** to open the VBA editor.

2. **Find the Currency Code**:
   - In the VBA editor, press **Ctrl + F** to open the Find dialog.
   - Search for the word **"dollars"** (or the currency that’s currently in the code).

3. **Replace the Currency**:
   - Replace **"dollars"** with the currency you want to use (e.g., "euros", "pounds", etc.).
   - Use **Replace All** to make the change throughout the code.

4. **Save the Workbook**:
   - When trying to save the workbook, you will encounter an error saying the workbook cannot be saved normally.
   - Simply click **No** when prompted.
   - You must save the file as a **macro-enabled workbook**. Save it as an **Excel Macro-Enabled Workbook (.xlsm)** to ensure the VBA code continues to work properly.

---

This custom formula provides a streamlined way to convert numbers into words directly within Excel, making it especially useful for reports, invoices, and other financial documents. By following the steps above, users can easily integrate and use the formula in their own Excel workbooks. Additionally, users can modify the currency by following the steps to replace the default "dollars" with any other currency, ensuring flexibility for different financial applications.
