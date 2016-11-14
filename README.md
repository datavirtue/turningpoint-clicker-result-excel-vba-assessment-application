An Excel spreadsheet application that accepts clicker poll and self-paced test result files, comparing them to show student progress.  

This is done weekly following this process:
1.	Students are assigned study materials (video/text)
2.	In class, questions are presented and answers collected using clickers (response devices)
3.	Students review study materials and do homework to cover weak areas
4.	Next class, students take a self-paced test using Turning Point
5.	Using the assessment and analysis spreadsheet application, both results are compared to show a graphical representation of student progress

What does the spreadsheet application do?

- imports the TurningPoint poll and self-paced test result files
- enables the teacher to map poll questions to one or many test questions
- computes a progress score for each student based on a comparison of poll results to test results 
- stores the progress results for each student in the specified workbook in the Assessments sheet and builds a graph for each week
	- the graph quickly shows the student and class formative assessment profile
- stores the test results for each student in the specified workbook in the tests worksheet and creates/updates a graph for the entire term
	- this graph quickly shows student and class summative performance

https://www.turningtechnologies.com/response-options
	

From the VBA code comments:
Used to assess students from weekly in-class polling and self-paced testing weekly and over
an entire term.  Output workbook contains a weekly improvement analysis of each student
as well as an aggregate chart showing test performance as the term progresses.

An Excel application that tabulates improvement/regression over sets of
TurningPoint Response solution result files.  The assessment is computed as a
percentage (positive or negative) of improvement from the first set of results (formative)
to the second set of results (summative).  The assessments can be collated by using
the same target workbook for an entire class term.  A graph of the improvemnt is generated
from the scores.  It is assumed that the same participant list is used for both
response sessions.  This was originally designed to support assessment of
improvement for self-paced tests after in-class polling that is conducted weekly.  Test
results are stored on a seperate worksheet from the assessment and the assessment
scores are highlighted according to the test results to paint a clear, quick, and data-driven
picture of student progress.


