# FPCounter
Automatic calculation of FP based on the code in an SCM

This tool allows you to count the function points of a project stored on a SCM.
Currently the tool works by connecting to one or more instances of MS Team Foundation Server, to one or more collection. 
For each collection it selects individual projects, then laws  Work Item custom "CI Software"  to get the name of the path
to be analyzed and writes the results in the custom work item named QAReport
