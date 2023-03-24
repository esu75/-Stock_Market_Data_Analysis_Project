Stock_Market_Data_Analysis_Project
----------------------------------
For this project VBA scripting is used to analyze generated stock market data
Goal:

Create a script that loops through all the stocks and outputs the following information:

      a. The ticker symbol
      b. Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.
      c. The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.
      d. The total stock volume of the stock. 
      e. Return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume". 
      f. Make the appropriate adjustments to the VBA script to enable it to run on every worksheet (i.e, years: 2018, 2019 and 2020) at once.
      
    step 1: Create the class for all attributes, declare and initialize varibales
    step 2: Create summary table headers
    step 3. Get the open price for the first ticker
    step 4: Iterate for each worksheet using 'for loop' through all the ticker names
    step 5: calculate the Yearly change using the close and open price  differences and update the table
    step 6: calculate the the percentage change from the opening price at the beginning of a given year to the closing price at the end of that year
    step 7: calculate the total volume from cells(i,7)
  summary tbale for steps 1 through 7:
  ------------------------------------

![image](https://user-images.githubusercontent.com/118146659/227642324-02bb9942-8f16-4772-b7ff-74d215bf3602.png)

  
