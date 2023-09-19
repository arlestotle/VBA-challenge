# VBA-challenge

In the first part of this challenge we were tasked to create a macro that generated a summary table of multiple year stock data with four new columns: Ticker, Yearly Change, Percent Change, and Total Stock Volume. The inital challenge of this assignment was to create a for loop that looks through all worksheets in the excel file and summarizes the data for each type of stock. Through the help of stack overflow I found a line of code that uses the operator <> to create the Tick Counter. 
The next major issue I faced in the first part of the assignment was creating output for the total stock volume column. I used google and found the function WorksheetFunction.Sum(Range("")) which allowed me to sum up the volume of stocks for each different ticker. I found it was necessary to reset the ticker counter by 1 inorder to find the sum of volumes for all different stocks. 
In the second half of the assignment we were tasked to create another summary table which returned the stocks with the greatest percent increase and decrease and the greatest total volume. I was able to generate this code through the use of for loops and conditionals. At the end, I formatted the summary table to include a percent sign on the value of stocks with the greatest increase and decrease and I changed the value of the greatest total volumn into scientific notation. 

Through the use of Stack Overflow, Github, the help of some of my peers I was able to find the code for:
  If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
   If ws.Cells(j, 3).Value <> 0 Then
      PercentChange = (ws.Cells(TickCounter, 11).Value / ws.Cells(j, 3).Value)
      ws.Cells(TickCounter, 12).Value = Format(PercentChange, "Percent")
  ws.Cells(TickCounter, 13).Value = WorksheetFunction.Sum(Range(ws.Cells(j, 7), ws.Cells(i, 7)))
    If ws.Cells(i, 12).Value > GreaterIncrease Then
        GreaterIncrease = ws.Cells(i, 12).Value
        ws.Cells(2, 17).Value = ws.Cells(i, 10).Value
       Else
        GreaterIncrease = GreaterIncrease
       End If

I was told by one of the TAs a .vbs file would not be necessary for the assignment submission and that a textfile of our VBS code would be sufficient. The text file containing the code of my VBA macro got a little messy when I loaded it onto github, so I went ahead and also attached the alphabet_testing.xlsm which includes the macro that created the output for my screenshots. 
