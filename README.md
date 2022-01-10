# stocks-analysis 

## 1. Overview of Project: 

**Data Analytics Bootcamp - Challenge Module 2**

	At the click of a button, analyze an entire dataset. 
	Select which dataset to analyze by typing the dataset name.
	Run analysis using excel macros and compare results from an optimized vba code and a non-optimized one.
	
### Datasets
Data of 12 different Wall-street stocks from 2017 and 2018

### Analysis
Get return of each stock to evaluate which one is more profitable

## 2. Results: 
+ **2017 Results**
	+ Refactored code
	![0.253 s](https://github.com/IrvingHdez/stocks-analysis/blob/main/Resources/VBA_Challenge_2017.PNG)
	+ Non Refactored code
	![0.1.355 s](https://github.com/IrvingHdez/stocks-analysis/blob/main/Resources/Results_2017_NonOptimized.PNG)
	+ **Optimized time = 1.102 s**
+ **2018 Results**
	+ Refactored code
	![0.316 s](https://github.com/IrvingHdez/stocks-analysis/blob/main/Resources/VBA_Challenge_2018.PNG)
	+ Non Refactored code
	![1.941 s](https://github.com/IrvingHdez/stocks-analysis/blob/main/Resources/Results_2018_NonOptimized.PNG)
	+ **Optimized time = 1.625 s**

## 3. Summary: 
### 3.1 What are the advantages or disadvantages of refactoring code?
There seems to be a clear processing time reduction of more than 1 sec which is great could be more noticible in larger datasets.

### 3.2 How do these pros and cons apply to refactoring the original VBA script?
The major time processing time reduction came from eliminating a nested loop.  
The number of iterations in the non refactored code was the number of records time number of tickers, 
while the optimized code only iterated once over the number of records.
