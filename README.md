# Excel-Creating-Link
Open the workbook and save a copy of it.

Select Open.
    Select Browse.
    open the file Develetech Sales.xlsx.

Paste a link to Total Sales.
    Verify that the Data worksheet is selected, and press Ctrl+End to view the last cell in the worksheet.
    Cell N2030 contains the total of all sales. For quick reference, it might be useful to show a copy of this information at the top of the worksheet.
    Press Ctrl+C to copy the cell.
    Press Ctrl+Home to return to the top left of the worksheet.
    Select cell B3.
    Select Home→Paste→Paste Link.
    View the value in cell B3 and its formula.
    The total sales value is now shown in cell B3. This is the simplest form of a linked cell. The formula provides a reference to another cell within the same worksheet. Although you used the Paste Link command to produce this formula, you could have simply typed the formula directly into the cell if you knew the cell you wanted to link.

Create links to the sales totals of three salespersons.
    Select cell B5 and type = to begin entering a formula.
    Select the Austin worksheet and select cell B3.
    Press Ctrl+Enter.
    Austin's sales value is now shown in cell B5.
    Note: You could have entered the formula by just pressing Enter. Using Ctrl+Enter leaves cell B5 selected, enabling you to see the formula without having to re-select the cell.
    Create similar links for Scott and Watson in cells B6 and B7, respectively.
    Verify the sales totals for each salesperson are displayed.

Create external links to summarize the quarterly workbook data.
    Select File→Open and browse to C:\091138Data\Working with Multiple Worksheets and Workbooks.
    Click once on Q1 Sales.xlsx to select it. Then, while pressing Ctrl, select Q2 Sales.xlsx, Q3 Sales.xlsx, and Q4 Sales.xlsx.
    The four workbook files are selected and added to the File name text box.
    Select Open to open all four workbooks.
    Using the taskbar, switch back to My Develetech Sales.xlsx.
    Select the Summary worksheet. Verify that cell B3 is selected, and type =
    Switch to the Q1 Sales.xlsx workbook, select cell D22, and press Ctrl+Enter.
    Verify the Quarter 1 sales are displayed on the Summary worksheet.
    Select cell B4 and type =
    Switch to the Q2 Sales.xlsx workbook, select cell D22, and press Ctrl+Enter.
    Note: You may have noticed that you are not using AutoFill to copy the formula, and you cannot. This is because you are changing workbooks, and each worksheet name corresponds to the quarter.
    Select cell B5 and type =
    Switch to the Q3 Sales.xlsx workbook and select cell D22 and press Enter.
    The Quarter 3 sales amount is linked to cell B5, and cell B6 is then selected.
    Select cell B6 and type =
    Switch to the Q4 Sales.xlsx workbook and select cell D22 and press Enter.
    Verify the quarterly sales values appear as shown here.

Arrange the windows of the open files so you can see all five simultaneously.
    Select View→Arrange All.
    Verify the Tiled radio button is selected and select OK.
   
View a linked cell update when you change the cell it is linked to.
    In the Q1 Sales.xlsx workbook, select cell B2.
    Type 1000 and press Enter.
    Verify that cell B3 in My Develetech Sales.xlsx changes automatically to $34,670.
