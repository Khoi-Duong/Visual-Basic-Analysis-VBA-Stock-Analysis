# The VBA of Wall Street

### Background
You are well on your way to becoming a programmer and Excel expert! In this homework assignment, you will use VBA scripting to analyze generated stock market data.

### Before You Begin
1. Create a new repository for this project called VBA-challenge. Do not add this assignment to an existing repository.

2. Inside the new repository that you just created, add any VBA files that you use for this assignment. These will be the main scripts to run for each analysis.

### Files
Download the following files to help you get started:

https://static.bc-edx.com/data/dl-1-2/m2/lms/starter/Starter_Code.zip

### Instructions
Create a script that loops through all the stocks for one year and outputs the following information:

* The ticker symbol

* Yearly change from the opening price at the beginning of a given year to the closing price at the end of that year.

* The percentage change from the opening price at the beginning of a given year to the closing price at the end of that year.

* The total stock volume of the stock. The result should match the following image:

![image](https://user-images.githubusercontent.com/119692456/235330925-c30ace82-26f3-47ea-9524-971dd0a1af27.png)

* Add functionality to your script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume". The solution should match the following image:

![image](https://user-images.githubusercontent.com/119692456/235330931-eb14c68c-f72a-43b8-a3cb-95c00be353d3.png)

* Make the appropriate adjustments to your VBA script to enable it to run on every worksheet (that is, every year) at once.

NOTE : Make sure to use conditional formatting that will highlight positive change in green and negative change in red.

# Other Considerations
* Use the sheet alphabetical_testing.xlsx while developing your code. This dataset is smaller and will allow you to test faster. Your code should run on this file in under 3 to 5 minutes.

* Make sure that the script acts the same on every sheet. The joy of VBA is that it takes the tediousness out of repetitive tasks with the click of a button.
