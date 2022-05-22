# Stock-Analysis
Overview of Project: Explain the purpose of this analysis.  
 In this project, we are refactoring code to speed up our analysis of stock data.  The data spans 
 12 tickers over 2 years.  There are approximately 3000 lines of data for each year.  
 With the initial code, VBA would run through the ~3000 line of data 12 times, for a total of ~36,000 runs.
 Our goal is to find a more direct route to read all of the lines and gather more information with less passes.
 
Results:
With the initial code, we were asking the user to input a year, and the code would go though the data line by line, first comparing ticker to the line above and below
We added more lines of code to tell VBA

Running the initial code, the time for each year was as follows:  
2017: ![Init_VBA_Challeng_2017](https://user-images.githubusercontent.com/103051630/169701514-b42afa8c-106a-4cdf-a42b-32528331de73.png)  
2018: ![Init_VBA_Challenge_2018](https://user-images.githubusercontent.com/103051630/169701517-ddfccde9-113f-4b99-8583-64385f715d53.png)

After refactoring the code, the time for each year was as follows:  
2017: ![Refac_VBA_Challenge_2017](https://user-images.githubusercontent.com/103051630/169701617-18a95ed9-94b1-4349-b30a-116fb5fca82c.png)  
2018: ![Refac_VBA_Challenge_2018](https://user-images.githubusercontent.com/103051630/169701633-d05a7383-baba-47ac-b791-ca0b013f5a86.png)




Summary: In a summary statement, address the following questions.
What are the advantages or disadvantages of refactoring code?
How do these pros and cons apply to refactoring the original VBA script?
