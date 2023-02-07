# Maintainability Index Calculator
## Introduction

This python script is designed to compute the average maintainability index of projects from the file generated by Visual Studio. The maintainability index is a measure of how maintainable the source code of a project is, based on various metrics such as code complexity, duplication, and coding style.

To run this project, you will need:

  - Python 3.x installed on your local machine
  - The pandas and numpy Python packages installed
  - A local installation of Excel or equivalent software that can read and write Excel files

## Setup

  Clone or download the project repository to your local machine.  
  Navigate to the project directory in your terminal or command prompt  
  
  Install the required Python packages by running the following command:
`
pip install pandas numpy
`
## Usage
  Copy the Excel file containing the project data into the project directory [^1]
  
  Run the following command in your terminal or command prompt to start the project:
```
python MI_Compute.py
```

  The project will filter the data in the Excel file and output the average maintainability index for each project, as well as the overall average, total count of projects, total count of projects with NaN value, and total count of projects excluding projects with NaN value.
  The final result will be saved to two CSV files: `result.csv` and `maintainability_index_values_only.csv`  
  
  **results.csv** would contain the MI value of a project and the specific project name.  
  **maintainability_index_values_only.csv** would contain the MI value of the project only to be easily copied as row data to be transfer to other spreadsheet applications.

## Customization
### XLSX File Name
___
For the xlsx file name you could edit this line of code to match what you have: https://github.com/PatrickNulla/MI-Calculator/blob/d27a2500b1e383426c3d49148e6b42454c15acd9/MI_Compute.py#L6  
### Filter JSON
___
If you want to exclude projects from the computation, simply add their names to the filter.json file and re-run the project: https://github.com/PatrickNulla/MI-Calculator/blob/23dcc0365beecfd0a73d1fde3576c70fe0bb24c2/filter.json#L1-L7  
#### Sample JSON:  
```json
{
    "filteredProjects":[
        "FolderName1\\FolderName4\\SolutionName.ProjectName1",
        "FolderName2\\FolderName5\\SolutionName.ProjectName2",
        "FolderName3\\FolderName6\\SolutionName.ProjectName3"
    ]
}
```
You could obtain this also on the console output of the script in this format: Scope - **ProjectName** : MI (What you want is the ProjectName)  
### Scope JSON
___
If you want to add more scope, simply add their scope values to the scope.json file and re-run the project: https://github.com/PatrickNulla/MI-Calculator/blob/23dcc0365beecfd0a73d1fde3576c70fe0bb24c2/scope.json#L1-L7
#### Sample JSON:  
```json
{
    "columnScope":[
        "Assembly",
        "Namespace",
        "Type",
        "Member"
    ]
}
```  

[^1]: Open Visual Studio Solution -> Analyze -> Calculate Code Metrics -> For Solution -> (On Code Metrics Result) Open List in Microsoft Excel -> Save File to Project Directory.
