# Convert Fortify SCA Developer Workbook XML Report to Excel

## Summary
This is a beta version tool that could aid to generate Excel report from Fortify SCA Developer Workbook XML report using Python v3.

## Install required packages
pip install -r requirements.txt

## Steps

 1. Scan your project and open the .fpr file with Fortify Audit Workbench
 2. Navigate to **Tools > Reports > Generate Legacy Report**
 3. Choose **Fortify Developer Workbook**
 4. Choose Save report 
 5. Choose XML format and save
 6. Open the **generate-excel** script and **edit** the filename variable to point to your report generated on step 5
 7. Run the script
