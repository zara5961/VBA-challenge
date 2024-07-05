# VBA-challenge

## Background
You are well on your way to becoming a programmer and Excel expert! In this homework assignment, you will use VBA scripting to analyze generated stock market data.

## Before You Begin
Create a new repository for this project called VBA-challenge. Do not add this assignment to an existing repository.

Inside the new repository that you just created, add any VBA files that you use for this assignment. These will be the main scripts to run for each analysis.

### Files
Resources/Multiple_year_stock_data.xlsx

### Instructions
Create a script that loops through all the stocks for each quarter and outputs the following information:

* The ticker symbol

* Quarterly change from the opening price at the beginning of a given quarter to the closing price at the end of that quarter.

* The percentage change from the opening price at the beginning of a given quarter to the closing price at the end of that quarter.

* The total stock volume of the stock. The result should match the following image:

Moderate solution
* Add functionality to your script to return the stock with the "Greatest % increase", "Greatest % decrease", and "Greatest total volume". The solution should match the following image:

Hard solution
* Make the appropriate adjustments to your VBA script to enable it to run on every worksheet (that is, every quarter) at once.

  ### NOTE
Make sure to use conditional formatting that will highlight positive change in green and negative change in red.

### Other Considerations
* Use the sheet alphabetical_testing.xlsx while developing your code. This dataset is smaller and will allow you to test faster. Your code should run on this file in just a few seconds.

Make sure that the script acts the same on every sheet. The joy of VBA is that it takes the tediousness out of repetitive tasks with the click of a button.

Requirements
Retrieval of Data (20 points)
The script loops through one quarter of stock data and reads/ stores all of the following values from each row:

ticker symbol (5 points)

volume of stock (5 points)

open price (5 points)

close price (5 points)

Column Creation (10 points)
On the same worksheet as the raw data, or on a new worksheet all columns were correctly created for:

ticker symbol (2.5 points)

total stock volume (2.5 points)

quarterly change ($) (2.5 points)

percent change (2.5 points)

Conditional Formatting (20 points)
Conditional formatting is applied correctly and appropriately to the quarterly change column (10 points)

Conditional formatting is applied correctly and appropriately to the percent change column (10 points)

Calculated Values (15 points)
All three of the following values are calculated correctly and displayed in the output:

Greatest % Increase (5 points)

Greatest % Decrease (5 points)

Greatest Total Volume (5 points)

Looping Across Worksheet (20 points)
The VBA script can run on all sheets successfully.
GitHub/GitLab Submission (15 points)
All three of the following are uploaded to GitHub/GitLab:

Screenshots of the results (5 points)

Separate VBA script files (5 points)

README file (5 points)

Grading
This assignment will be evaluated against the requirements and assigned a grade according to the following table:

Grade	Points
A (+/-)	90+
B (+/-)	80–89
C (+/-)	70–79
D (+/-)	60–69
F (+/-)	< 60
Submission
To submit your Challenge assignment, click Submit, and then provide the URL of your GitHub repository for grading.

NOTE
You are allowed to miss up to two Challenge assignments and still earn your certificate. If you complete all Challenge assignments, your lowest two grades will be dropped. If you wish to skip this assignment, click Next, and proceed to the next module.

Comments are disabled for graded submissions in Bootcamp Spot. If you have questions about your feedback, please notify your instructional staff or your Student Success Advisor. If you would like to resubmit your work for an additional review, you can use the Resubmit Assignment button to upload new links. You may resubmit up to three times for a total of four submissions.

IMPORTANT
It is your responsibility to include a note in the README section of your repo specifying code source and its location within your repo. This applies if you have worked with a peer on an assignment, used code in which you did not author or create sourced from a forum such as Stack Overflow, or you received code outside curriculum content from support staff such as an Instructor, TA, Tutor, or Learning Assistant. This will provide visibility to grading staff of your circumstance in order to avoid flagging your work as plagiarized.

If you are struggling with a challenge assignment or any aspect of the academic curriculum, please remember that there are student support services available for you:

Ask the class Slack channel/peer support.

AskBCS Learning Assistants exists in your class Slack application.

Office hours facilitated by your instructional staff before and after each class session.

Tutoring GuidelinesLinks to an external site. - schedule a tutor session in the Tutor Sessions section of Bootcampspot - Canvas

If the above resources are not applicable and you have a need, please reach out to a member of your instructional team, your Student Success Advisor, or submit a support ticket in the Student Support section of your BCS application.

References
Data for this dataset was generated by edX Boot Camps LLC, and is intended for educational purposes only.

