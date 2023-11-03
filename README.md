# python_diaries 

### A series where I detail my experiences, thoughts, and struggles learning Python as an R user.

## Python Diaries 1 -- Large Pumpkins :jack_o_lantern:

In this first installment of the Python Diaries, I take a look at a global dataset from The Great Pumpkin Commonwealth on large pumpkins. 

### My goal for this project:
Get a feel for how Python works and do a little bit of exploratory data analysis in a low-stakes environment. 

### Skills I practiced in this project:
- Creating a Jupyter Notebook 
- Importing a .csv file
- Installing and importing packages (Pandas, Numpy, Matplotlib, Seaborn) 
- Cleaning data 
- Manipulating data
- Summarizing data
- Visualizing data 
  - Histogram
  - Bar chart 
  - Box plot 


### General reflections: 
As primarily an R user, there were some differences that stood out to me after getting to know Python for a little bit. Firstly, with R there is primarily one IDE that manages and executes R code, which is RStudio. While there are a few other alternatives for RStudio, they don't compare to the number of IDEs, text editors, and notebooks available for Python code. I decided to run my Python code in Jupyter Notebook because I read that it is commonly used in data science and also proof of knowledge exercises. However, I plan on testing out other IDEs, etc. to manage and execute Python code. 

In general, the language differences between R and Python is akin to the differences between Spanish and Portuguese. The languages are close enough whereas I can get by in Python and understand the fundamentals of what I am doing, but I would still need to refer to guides in order to execute the code correctly. There are also smaller "grammatical" differences between the two languages where R uses "<-" as the assignment operator and Python uses "=". 

Lastly, it seems that from my research that people tend to say that R is "intuitive" and best for "statistical analyses" and "data science" while Python is best for "production" and it will give you real "programming skills". From the get-go, R seemed more intuitive to me in the beginning than Python has so far. However, I do not want to make any hard-core judgment statements until I have some more experience with Python using it in different contexts, i.e., visualizations, statistics, machine learning, etc. 

### What I want to learn for next time:
It would be really handy to learn how to add comments to code in Juptyer Notebook. While this first notebook is a bit messy without comments, I hope that my goal and intentions in this elementary stage of learning Python still comes through. 


### Available Files: 
### [Python Diaries 1 Notebook](https://github.com/emartin43/python_diaries/blob/e0fd13e67e098e6e0b5574d65463545fbf96ec95/pd1_large_pumpkins/python_diaries_1_script.ipynb)
### [Pumpkins Data](https://github.com/emartin43/python_diaries/blob/e0fd13e67e098e6e0b5574d65463545fbf96ec95/pd1_large_pumpkins/pumpkins.csv)


## Python Diaries 2 -- Solitary Confinement in Detention Facilities 

In this second installment of the Python Diaries, I analyze solitary confinement data in detention facilities from the International Consortium of Investigative Journalists (ICIJ) report "Solitary Voices". 

### My goal for this project:
Practice the Python skills I was introduced to in Python Diaries #1 and feel more comfortable with them.

### Skills I practiced in this project:
- Creating a Jupyter Notebook 
- Importing a .csv file
- Installing and importing packages (Pandas, Numpy, Matplotlib, Seaborn) 
- Cleaning data 
- Manipulating data
- Summarizing data
- Visualizing data 
  - Histogram
  - Bar chart 
  - Box plot 
- Creating comments 
- Setting data display formatting 
- Subsetting Data 


### General reflections: 
After the last Python Diaries, I wanted to immediately jump into geospatial analysis or linear regressions, but I realized that it would be better to become more comfortable with basic Python concepts first before moving on to something more complicated. 

I was so excited to dive into the world of data science and learn all the cool libraries and techniques. But I quickly realized that before I could even think about using those libraries, I needed to have a strong understanding of the basics. When it comes to data science, just like many STEM fields, understanding the basics is like the foundation of a house. If the foundation is weak, the rest of the house will crumble. But with a strong foundation, it's possible to build anything. 

It was really rewarding to work with this dataset because of all the potential ways it can be used to help people and communities. Data science can be a powerful tool for promoting social good, especially when it comes to analyzing complex issues such as solitary confinement in detention facilities. By using techniques like machine learning and statistical analysis, data scientists can uncover patterns and insights that might otherwise go unnoticed. For example, in the context of solitary confinement, data science can be used to analyze large amounts of data on the conditions and outcomes of this practice, such as the length of time individuals spend in isolation, the impact on their mental health, and the correlation between solitary confinement and recidivism. This information can be used to inform policy and advocacy efforts aimed at reducing the use of solitary confinement and improving conditions for those who are incarcerated. Additionally, data science can also be used to monitor and evaluate the effectiveness of these efforts over time.


### What I want to learn for next time:
Next time, I want to take my Python skills to the next level by tinkering with geospatial data or running simple linear regressions. 


### Available Files: 
### [Python Diaries 2 Notebook](https://github.com/emartin43/python_diaries/blob/ec2da0811c093c9b4d68a7a38d674fd4c2b63967/pd2_solitary_confinement/Python%20Diaries%20%232.ipynb)
### [Solitary Confinement Data](https://github.com/emartin43/python_diaries/blob/ec2da0811c093c9b4d68a7a38d674fd4c2b63967/pd2_solitary_confinement/icij-solitary-voices-final-dataset-for-publication.csv)



## Python Diaries 3 -- Medical Students Mental Health

In the third installment of the Python Diaries, I analyze medical students' mental health in Switzerland. This dataset explores medical students' empathy, mental health, and burnout in Switzerland (Data Carrard et al. 2022 MedTeach.csv). It compiles important demographic information as well as self-reported data and results from psychological tests to give a comprehensive picture of the mental states of students in the medical field. 

### My goal for this project:
Practice more visualization techniques as well as delve into basic one-variable linear regressions.

### Skills I practiced in this project:
- Creating a Jupyter Notebook 
- Importing a .csv file
- Installing and importing packages (Pandas, Numpy, Matplotlib, Seaborn) 
- Cleaning data 
  - Checking for outliers
- Manipulating data
- Summarizing data
  - Correlation Matrices
- Visualizing data 
  - Histogram
  - Bar chart 
  - Box plot 
  - Linear Regression
  - Heatmap
  - Stripplot
- Creating comments 


### General reflections: 
It was really neat to be able to create one-variable linear regressions from the data. With the other installments of the Python Diaries, I was focusing more on the EDA aspect of data wrangling, which is very important, but it was also nice to transition onto more predictive analytics. 

We run a linear regression of depression on anxiety to see if they were associated at all for these medical students. We find that there was a 0.72 correlation between the self-reported survey results of depression and anxiety in these medical students, indicating that these two mental illnesses may be comorbid, which aligns with the existing literature on common comorbidities (AL-Asadi et al., 2015). Additionally, we find a moderate correlation (0.61) between self-reported depression scores and self-reported exhaustion of medical students. Surprisingly, we find a little to no correlation between self-reported depression and hours of study per week of the participant, suggesting that time spent studying does not impact depression and depression (and correlated exhaustion) does not decrease how much these medical students study. 

The aim with this analysis was to enhance the comprehension of the impact of being a medical student on one's health and wellbeing. By examining the specific factors that may influence various outcomes, we can improve educational systems, benefiting both medical students and their future patients.


### What I want to learn for next time:
Next time, it would be interesting to work on web scraping or pulling data using an API.  


### Available Files: 
### [Python Diaries 3 Notebook](https://github.com/emartin43/python_diaries/blob/08ccc5dd24565ae61540ff6a6387974210dfa05a/pd3_mentalhealth/python_diaries_3.ipynb)
### [Medical Student Mental Health Data](https://github.com/emartin43/python_diaries/blob/08ccc5dd24565ae61540ff6a6387974210dfa05a/pd3_mentalhealth/med_mental_health.csv)
### [Codebook](https://github.com/emartin43/python_diaries/blob/08ccc5dd24565ae61540ff6a6387974210dfa05a/pd3_mentalhealth/Codebook%20Carrard%20et%20al.%202022%20MedTeach.csv)


## Python Diaries 4 -- Data Extraction of Bonneville Power Administration Lighting Calculator Files

This is a project I worked on during my internship at Tacoma Public Utilities/Tacoma Power in the Customer Energy Programs group working on energy efficiency and conservation. 

### My goal for this project:
Develop an automatic data extraction process to scrape information from disparate Excel spreadsheets, join the data into one data frame, transform the data, and load the data into a Snowflake cloud database as part of the ETL process in data warehousing. 

### Skills I practiced in this project:
- File handling with the os module
- Extracting data from Excel files using openpyxl and Pandas
- Data cleaning and transformation
- Regular expression (regex) usage for data parsing
- Iteration and looping through files and directories
- Data conversion and formatting
- Pandas DataFrame manipulation
- Error handling with None values
- Concatenation and combination of DataFrames
- Documentation and comments for code clarity
- Quality assurance practices to ensure data accuracy and integrity

### General reflections:
When I first received this request, I wasn't entirely sure how to approach it. I had a little bit of familiarity with scraping PDF files, though that was a major hassle, and that was when I first started learning how to code. Now that I'm more confident in my programming abilities, it's nice to have experience working through tough assignments. 

After doing some research on what strategy to take, I spent a lot of time just thinking and reflecting on the best way to approach my particular task. I like externalizing my thought processes because it helps me clarify and organize all the different ideas I might have to tackle a project, which is why my whiteboard and whiteboard markers are my favorite office supplies.

One of the biggest life lessons I've learned from programming is breaking up one big problem into many smaller problems. It definitely is so helpful and calming to know that I can always break up any problem I have into smaller, more manageable tasks. For example, in this project, before I could even think about scraping data, I had to write a chunk of code that converted all the data files to .xlsx because the package 'openpxl' only works with .xlsx files. Then, I ran through the logic of scraping one single file before writing the loop to extract data from multiple files in multiple subfolders. There were many instances of having to break my logic down to solve a simpler problem. 

### What I want to learn for next time:
The project I'm currently working on is another data extract of bi-yearly reports. It would be neat to take this project to the next step by taking the data uploaded into Snowflake and visualizing any insights in Tableau, which is a software I'm currently studying to certify for. 

### Available Files:
### [Python Diaries 4 Code](https://github.com/emartin43/python_diaries/blob/main/pd4_extraction/LC_Extraction.py)
