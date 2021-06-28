###### stock-analysis-VBA-challenge 

## Overview of Project 

Steve is analyzing stocks for his parents so they can decide which stocks to invest in. They thought they wanted to invest in DQ but the analysis showed that perhaps there are more profitable and stable stocks. To understand which stocks are best, Steve asked for some help analyzing stock performance over the last few years for all stocks. This analysis allowed Steve to do just that at the click of a button in the Excel sheet. 

## Results 
Results: Using images and examples of your code, compare the stock performance between 2017 and 2018, as well as the execution times of the original script and the refactored script.
#### 2017 vs 2018 
The analysis shows that the stocks generally fared better in 2017, except for TERP which was down 7.2% that year and recovered slightly in 2018 to be down only 5%. In 2018, ENPH and RUN were the only 2 stock, of the 12, that gained rather than lost value over the year. 

<img width="258" alt="2017 Stock Performance" src="https://user-images.githubusercontent.com/85316096/123570287-249fb280-d785-11eb-9a73-ba5fb84ad0b3.png"> <img width="261" alt="2018 Stock Performance" src="https://user-images.githubusercontent.com/85316096/123570303-2b2e2a00-d785-11eb-8d85-1463e02b5ee1.png">

#### Execution Times Between Original and Refactored Scripts 
The execution time of the original scripts for both 2017 and 2018 was 0.5546875 seconds. 

<img width="274" alt="PreRefactored Time 2017" src="https://user-images.githubusercontent.com/85316096/123570718-0dad9000-d786-11eb-9c38-9d8e2e435281.png"> <img width="267" alt="PreRefactored Time 2018" src="https://user-images.githubusercontent.com/85316096/123570704-071f1880-d786-11eb-966e-c38fb632b238.png">

The execution time after refactoring the code was quite a bit faster at 0.109375 and 0.1015625.

<img width="261" alt="VBA_Challenge_2017" src="https://user-images.githubusercontent.com/85316096/123570523-9bd54680-d785-11eb-893f-337d02e54da3.png"> <img width="260" alt="VBA_Challenge_2018" src="https://user-images.githubusercontent.com/85316096/123570528-a099fa80-d785-11eb-9e1d-5a6eb5609ee6.png">

## Summary

#### What are the advantages or disadvantages of refactoring code?
Refactoring code can speed up the analysis time, like it did in this project, and it can also make code easier to read/understand, follow a more logical construct and make the code easier to maintain. However, refactoring can also introduce bugs to the code and it can be time consuming. 

#### How do these pros and cons apply to refactoring the original VBA script?
In this project, refactoring the code made the program run faster to compile the results of a larger dataset, but it also meant creating a script that was harder to execute than the first time around. If I were a more experienced coder this may not have taken as much time, but it was difficult to debug and ensure that my code was pulling from the rights sheets and not encountering overflow errors. 
