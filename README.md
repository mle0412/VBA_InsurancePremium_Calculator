# Automated Insurance Premium Calculator

# Project summary:
A system that calculates the Term Life Insurance's premiums for the users, based on their input information.
Inputs are included:
  1. Name, email, and phone number. (for reporting purposes)
  2. Gender, age, and annual.  (for calculations)
  3. Kind and desired benefit amounts: (for calculation)
    a. Flat benefit: the beneficiary will receive a fixed amount.
    b. Percentage of annual income: the beneficiary will receive an amount that equates to the desired percentage of the insured's current annual income.
  4. Term of the insurance policy. (for calculation)
  5. Providers. (for calculation)
  6. Other behaviors: smoking, drinking (other factors such as work out or not, type of sugar consuming, ... etc. will be added in later)

# Features:
  1. Receive and input information of the user into Excel worksheet, which runs in the back.
  2. Plug those information to the calculator to calculate the insurance premium.
  3. Updating the fixed factors (such as Mortality rate, fee of providers) from the database. (In this project, the database is just another Excel workbook. Will be upgraded so it can retrieve data from a SQL server soon)
  4. Run the calculation in the back to calculate the Total Required Premium and the Monthly Payment for the user.
  5. Report multiple results to another worksheet, then sort the results by Premiums in ascending order.
  6. Create a graph that helps the user compare and choose the provider that best fits his/her needs.
  7. Automatically save the report as pdf file into the folder, which used to store this Excel workbook.
  8. Automatically send an email to the email address provided by the user.
  *** To be added in soon: send the information of the user (with notice and consent) to the database system for further analysis.

# How to use:
  1. Input all of the required information (Name, age, desired benefits, ... etc. )
  2. Click on Update button to get the most recent factors from the data sources then choose the providers.
  3. Click on Premium Calculation button for running the calculator. (Calculator will be run in the back, so nothing will happen)
  4. Click on Premium Comparision button for results to be reported and sorted. (Report will be run in the back, so nothing will happen)
  5. Choose how Report will be exported (save as a PDF or send to the email provided) and click on Print/Email button to execute.
  *** Note: because the VBA code execution is suspended while the user is in Print Preview, the Preview mode will not be available at this time.
  
# Note:
  1. Excel VBA Macro must be enabled.
  2. Microsoft Outlook Object Library in References must be enabled.
  3. Both the Insurance_Premium and Data_Sources file must be saved in: ```D:\Personal Learning\VBA ```

# Disclaimer:
  1. The quotes which reported from this Calculator will not reflect the true premium that user will be required in real-life.
  2. Many of the data is made up (such as the Name and Fee of providers, Probability that Smoking and Drinking will increase the mortality, ... etc.)

# License:
[GNU GPLv3] (https://choosealicense.com/licenses/gpl-3.0/)
