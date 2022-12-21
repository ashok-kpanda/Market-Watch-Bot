*** Settings ***
Documentation       Dispatcher Market Watch

Library             RPA.Robocorp.WorkItems
Library             Collections
Library             RPA.Excel.Files
Library             RPA.Tables


*** Variables ***
${Symbol}       SYMBOL


*** Tasks ***
Dispatcher Market Watch
    TRY
        ${sales_reps}=    Read Input Excel
    EXCEPT
        Log    Unable to read Excel
    END
    FOR    ${Item}    IN    @{sales_reps}
        Create Output Work Item    ${Item}
        Save Work Item
    END


*** Keywords ***
Read Input Excel
    Open Workbook    AA500_Input.xlsm
    Read Worksheet    Sheet1
    ${sales_reps}=    Read Worksheet As Table    header=True
    RETURN    ${sales_reps}
