# Stock-Analysis
* performed a stock-analysis with VBA

## Overview of Project
* Steve obtained a list of stock market in year 2017 and 2018, and we want to create a Macro that can perform stock analysis on entire stock market over a year for steve. 

### Purpose
* The purpose of this project is to create Macro can generate a report that shows an analysis in total daily voumne in a year and annual return percentage for each tickers. We also wanted to refactored the code so it runs much faster because Steve can run the analysis faster if he has a longer list of stock market. 

## Result
### Stock Performance Between 2017 and 2018
* In 2017, majority of the stocks on the list will have a positive return in investment. The stock ticker DQ and SEDG performed outstanding in 2017, with a 199.4% annual return for DQ and 184.5% annual return for SEDG. In current year, we noted the total daily volume does not have much connection with the annual return.

  <img width="365" alt="image" src="https://user-images.githubusercontent.com/107168891/177009556-78dff7bc-d45a-41b1-bc82-d454d52bff25.png">

* In 2018, majority of the stocks on the list will have a negative return in investment. Steven has to be careful with his investment. There are only two ticker that have a postive return, which is ENPH with a 81.9% of return and RUN with a 84% of return. In current year, we noted higher total dily volumne might be a factore for stocks to have a positive return rate. 
 
  <img width="365" alt="image" src="https://user-images.githubusercontent.com/107168891/177009657-5aca1780-db3e-4eb8-a35c-fc6de61864ac.png">
### Execute time Original Code vs. Refactored Code
* In the original code the execute time for both year 2017 and 2018 are around 0.4 seconds.

  <img width="252" alt="image" src="https://user-images.githubusercontent.com/107168891/177009958-d68deef2-fdfd-48e6-8ce2-69ea3d3f9dfa.png"> <img width="252" alt="image" src="https://user-images.githubusercontent.com/107168891/177009970-752d591e-027f-417e-a8f7-147211c4a1bb.png">

* After the code was refactored, we have an significant increase in excute time, with an average of 0.12 seconds.

  <img width="255" alt="image" src="https://user-images.githubusercontent.com/107168891/177010087-79c8a7ed-89c6-4f7b-837e-858bd5455cb6.png"> <img width="247" alt="image" src="https://user-images.githubusercontent.com/107168891/177010092-fca302fd-1c38-4dba-8c46-34f8348b1d79.png">

### Difference between Original Code and Refactored Code
* In the orginial code, we are using variables to capture our data and used nested for loop to run the analysis, which will take more times to run the analysis. 
  <img width="407" alt="image" src="https://user-images.githubusercontent.com/107168891/177010221-4af2b922-f1d6-49f7-8aaf-c57ac2b9bd0e.png">

* In the reactored code, we used array to capture our data and used multiple one for loop to run our data.

  <img width="371" alt="image" src="https://user-images.githubusercontent.com/107168891/177010251-07b2ed08-b098-4801-9f25-fab3b0669c7e.png">

## Summary
What are the advantages or disadvantages of refactoring code? 
* Some advantage of using refactored code is to run the data faster because it takes less memory. The code it self is much more simple with one for loop. 
* Some disvanage with refactored code is time consuming because it takes time to refactor the code. However, this statment might change if we have a larger set of data. 

How do these pros and cons apply to refactoring the original VBA script?
* Removing nested for loop can take less memory and by the look of it is much simple and easier to understand.
