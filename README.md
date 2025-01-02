'*****************************************************************************************************
' Script Purpose: Fetch data from the Data USA API to analyze changes in the United States population.
'
' Input:          Number of years (default is 10).
'
' Output:         Adds the following data to the first sheet:
'                 - Column A: Year
'                 - Column B: Population
'                 - Column C: Percent Increase
'                 Overwrites existing data in these columns.
'
' Requirements:   1. Add JsonConverter to the current project's Modules.
'                 - Source: https://github.com/VBA-tools/VBA-JSON
'                 2. Enable the following references from Tools -> References:
'                  - Microsoft WinHTTP Services
'                  - Microsoft Scripting Runtime
'
' Dependencies:   Active internet connection (to call the Data USA API).
'
' Author:         Kadi Jairus
' Date:           January 1, 2025
' Version:        1.0
'*****************************************************************************************************

