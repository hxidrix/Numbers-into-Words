 # Custom Formula for Converting Numbers into Words in Excel

This project involves creating a custom formula in Excel that converts numeric values into written words, such as "1000" being converted into "One Thousand". The formula is developed using VBA (Visual Basic for Applications) code, and it simplifies the process of displaying numbers as words within Excel. Here's how you can implement and use the custom formula:

**Method:**

1. **Copy the VBA Code**: 
   - First, copy the VBA code provided for converting numbers into words.

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

**Note:** The file contains macros, so it should only be opened in a macro-enabled Excel format (.xlsm) for all functionalities, including automated calculations and payslip generation, to work properly.
