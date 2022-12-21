*** Settings ***
Documentation       Performer Market Watch

Library             RPA.Browser.Selenium    auto_close=${False}
Library             RPA.Excel.Files
Library             RPA.Tables
Library             RPA.Windows    auto_close=${False}
Library             RPA.Desktop
Library             String
Library             Collections
Library             RPA.FileSystem
Library             RPA.Robocorp.WorkItems


*** Variables ***
${url}          https://www.marketwatch.com/
${sspath}       image_files${/}


*** Tasks ***
Performer Market Watch
    Set Local Variable    ${c}    2
    Set Local Variable    ${Exception}    NA
    Set Local Variable    ${nocompany}    No Company Present
    ${image_folderExist}=    Does Directory Exist    ${sspath}
    IF    ${image_folderExist} == True
        Empty Directory    ${sspath}
    ELSE
        Create Directory    ${sspath}
    END
    TRY
        ${Comp_Symbol}=    For Each Input Work Item    Reading Work Item
        TRY
            Open Market watch website
            Search all symbol in market watch and fetch data    ${Comp_Symbol}    ${c}    ${Exception}    ${nocompany}
        EXCEPT
            Log    Unable To Open Market watch website
        END

        Click Home Page
        Close Browser
    EXCEPT
        Log    Unable to get company Symbol from Workitem
    END


*** Keywords ***
Reading Work Item
    ${payload}=    Get Work Item Payload
    FOR    ${element}    IN    @{payload}
        IF    "${element}" == "SYMBOL"
            Log    ${payload}[${element}]
            ${Comp_Symbol}=    Set Variable    ${payload}[${element}]
            RETURN    ${Comp_Symbol}
        END
    END

Search all symbol in market watch and fetch data
    [Arguments]    ${Comp_Symbol}    ${c}    ${Exception}    ${nocompany}
    FOR    ${item}    IN    @{Comp_Symbol}
        TRY
            Log    ${item}
            ${elementexist}=    Click on search    ${item}
            IF    ${elementexist} == False
                ${companynamepresent}=    Click If Search Value Present
                IF    ${companynamepresent} == False
                    Click Home Page
                    Save Exception in Excel When wrong company keyword enter
                    ...    ${item}
                    ...    ${Exception}
                    ...    ${c}
                    ...    ${nocompany}
                    ${c}=    Evaluate    ${c} + 1
                    Sleep    1s
                ELSE
                    ${No_analyst_extractedlist}
                    ...    ${changepoint}=
                    ...    Extract the data IF no analyst tab and change value Present
                    ...    ${item}
                    Save excel If no analyst tab and change value table
                    ...    ${item}
                    ...    ${No_analyst_extractedlist}
                    ...    ${c}
                    ...    ${Exception}
                    Paste Graph in excel    ${item}    ${changepoint}
                    ${c}=    Evaluate    ${c} + 1
                END
            ELSE
                ${changevalue}=    Check Change Value Present or not
                IF    ${changevalue} == True
                    ${listexcel}    ${changepoint}=    Extract the Company Data From Market Watch    ${item}
                    save data into excel    ${item}    ${listexcel}    ${c}
                    Paste Graph in excel    ${item}    ${changepoint}
                    ${c}=    Evaluate    ${c} + 1
                ELSE
                    ${listexcel 1}
                    ...    ${changepoint}=
                    ...    Data Extraction If Anylist Tab Present and Change Value not Present
                    ...    ${item}
                    Add data to Excel If no Change value present
                    ...    ${listexcel 1}
                    ...    ${c}
                    ...    ${Exception}
                    ...    ${item}
                    Paste Graph in excel    ${item}    ${changepoint}
                    ${c}=    Evaluate    ${c} + 1
                END
            END
        EXCEPT
            Log    Unable to extract the data from Market Watch website
        END
    END

Open Market watch website
    Open Available Browser    ${url}
    Maximize Browser Window

Click on search
    [Arguments]    ${item}
    Click Element    css:button[class='btn btn--lighten btn--search j-btn-search']
    Click Element    css:button[class='btn btn--outline']
    Input Text    css:input[class='input input--search j-search-input']    ${item}
    Click Element    css:button[class='btn btn--secondary j-search-button']
    ${elementexist}=    Does Page Contain Element
    ...    css:[class='company__name']
    Log    ${elementexist}
    RETURN    ${elementexist}

Extract the Company Data From Market Watch
    [Arguments]    ${item}
    ${companyname}=    RPA.Browser.Selenium.Get Text    css:[class='company__name']
    ${value}=    RPA.Browser.Selenium.Get Text    css:[class='value']
    ${ChangePoint}=    RPA.Browser.Selenium.Get Text    css:[class='change--point--q']
    Log    ${ChangePoint}
    ${changepercent}=    RPA.Browser.Selenium.Get Text    css:[class='change--percent--q']
    Log    ${changepercent}
    ${percent}=    Remove String    ${changepercent}    %
    Log    ${percent}
    Sleep    5s
    Wait Until Keyword Succeeds    60x    0.5s    RPA.Desktop.Press Keys    Space
    Sleep    5s
    open application    calc.exe
    Wait Until Keyword Succeeds    60x    0.5s    RPA.Windows.Click    id:clearEntryButton
    Type Text    id:num${ChangePoint}Button
    RPA.Windows.Click    id:plusButton
    Type Text    id:num${percent}Button
    RPA.Windows.Click    id:equalButton
    ${result}=    Get Attribute    id:CalculatorResults    Name
    Log    ${result}
    Remove String    ${result}    Display is    null
    Log    ${result}
    Sleep    3s
    RPA.Windows.Click    id:Close
    ${finalresult}=    Replace String    ${result}    Display is    ${empty}
    Log    ${finalresult}
    ${Element color}=    Get Element Attribute    //*[@id="maincontent"]/div[2]/div[3]/div/div[2]/bg-quote    Class
    Log    ${Element color}
    ${Element color Final}=    Replace String    ${Element color}    intraday__change    ${empty}
    ${CloseValue}=    RPA.Browser.Selenium.Get Text    css:[class='table__cell u-semi']
    Log    ${CloseValue}
    ${changevalue}=    RPA.Browser.Selenium.Get Text
    ...    //*[@id="maincontent"]/div[2]/div[3]/div/div[4]/table/tbody/tr/td[2]
    Log    ${changeValue}
    ${changepercent1}=    RPA.Browser.Selenium.Get Text
    ...    //*[@id="maincontent"]/div[2]/div[3]/div/div[4]/table/tbody/tr/td[3]
    Log    ${changepercent1}
    ${percent1}=    Remove String    ${changepercent1}    %
    Log    ${percent1}
    Sleep    5s
    open application    calc.exe
    Wait Until Keyword Succeeds    60x    0.5s    RPA.Windows.Click    id:clearEntryButton
    Type Text    id:num${ChangeValue}Button
    RPA.Windows.Click    id:plusButton
    Type Text    id:num${percent1}Button
    RPA.Windows.Click    id:equalButton
    ${result1}=    Get Attribute    id:CalculatorResults    Name
    Log    ${result1}
    RPA.Windows.Click    id:Close
    ${finalresult1}=    Replace String    ${result1}    Display is    ${empty}
    Log    ${finalresult1}
    ${Element color1}=    Get Element Attribute
    ...    //*[@id="maincontent"]/div[2]/div[3]/div/div[4]/table/tbody/tr/td[2]
    ...    Class
    Log    ${Element color1}
    ${Element color Final 1}=    Replace String    ${Element color1}    table__cell    ${empty}
    Wait Until Keyword Succeeds    60x    0.5s    RPA.Desktop.Press Keys    Up
    Wait Until Keyword Succeeds    60x    0.5s    RPA.Desktop.Press Keys    Up
    #Wait Until Keyword Succeeds    40x    0.5s    Click Element    //*[@id="maincontent"]/div[2]/div[4]/mw-chart/label
    ${ss}=    RPA.Browser.Selenium.Screenshot
    ...    //*[@id="maincontent"]/div[2]/div[4]    ${sspath}${item}.jpg
    Sleep    5s
    ${OPEN}=    RPA.Browser.Selenium.Get Text    //*[@id="maincontent"]/div[6]/div[1]/div[1]/div/ul/li[1]/span[1]
    ${DAY RANGE}=    RPA.Browser.Selenium.Get Text
    ...    css:#maincontent > div.region.region--primary > div:nth-child(2) > div.group.group--elements.left > div > ul > li:nth-child(2) > span.primary
    ${52 WEEK RANGE}=    RPA.Browser.Selenium.Get Text
    ...    css:#maincontent > div.region.region--primary > div:nth-child(2) > div.group.group--elements.left > div > ul > li:nth-child(3) > span.primary
    ${MARKET CAP}=    RPA.Browser.Selenium.Get Text
    ...    css:#maincontent > div.region.region--primary > div:nth-child(2) > div.group.group--elements.left > div > ul > li:nth-child(4) > span.primary
    ${ SHARES OUTSTANDING}=    RPA.Browser.Selenium.Get Text
    ...    css:#maincontent > div.region.region--primary > div:nth-child(2) > div.group.group--elements.left > div > ul > li:nth-child(5) > span.primary
    ${PUBLIC FLOAT}=    RPA.Browser.Selenium.Get Text
    ...    css:#maincontent > div.region.region--primary > div:nth-child(2) > div.group.group--elements.left > div > ul > li:nth-child(6) > span.primary
    ${BETA}=    RPA.Browser.Selenium.Get Text
    ...    css:#maincontent > div.region.region--primary > div:nth-child(2) > div.group.group--elements.left > div > ul > li:nth-child(7) > span.primary
    ${REV. PER EMPLOYEE}=    RPA.Browser.Selenium.Get Text
    ...    css:#maincontent > div.region.region--primary > div:nth-child(2) > div.group.group--elements.left > div > ul > li:nth-child(8) > span.primary
    ${P/E RATIO}=    RPA.Browser.Selenium.Get Text
    ...    css:#maincontent > div.region.region--primary > div:nth-child(2) > div.group.group--elements.left > div > ul > li:nth-child(9) > span.primary
    ${EPS}=    RPA.Browser.Selenium.Get Text
    ...    css:#maincontent > div.region.region--primary > div:nth-child(2) > div.group.group--elements.left > div > ul > li:nth-child(10) > span.primary
    ${YIELD}=    RPA.Browser.Selenium.Get Text
    ...    css:#maincontent > div.region.region--primary > div:nth-child(2) > div.group.group--elements.left > div > ul > li:nth-child(11) > span.primary
    ${DIVIDEND}=    RPA.Browser.Selenium.Get Text
    ...    css:#maincontent > div.region.region--primary > div:nth-child(2) > div.group.group--elements.left > div > ul > li:nth-child(12) > span.primary
    ${EX-DIVIDEND DATE}=    RPA.Browser.Selenium.Get Text
    ...    css:#maincontent > div.region.region--primary > div:nth-child(2) > div.group.group--elements.left > div > ul > li:nth-child(13) > span.primary
    ${SHORT INTEREST}=    RPA.Browser.Selenium.Get Text
    ...    css:#maincontent > div.region.region--primary > div:nth-child(2) > div.group.group--elements.left > div > ul > li:nth-child(14) > span.primary
    ${% OF FLOAT SHORTED}=    RPA.Browser.Selenium.Get Text
    ...    css:#maincontent > div.region.region--primary > div:nth-child(2) > div.group.group--elements.left > div > ul > li:nth-child(15) > span.primary
    ${AVERAGE VOLUME}=    RPA.Browser.Selenium.Get Text
    ...    css:#maincontent > div.region.region--primary > div:nth-child(2) > div.group.group--elements.left > div > ul > li:nth-child(16) > span.primary
    Sleep    3s
    Wait Until Keyword Succeeds
    ...    20x
    ...    0.5s
    ...    Click Element
    ...    //*[@id="maincontent"]/div[5]/div/div/li[6]/a
    ${Rating}=    RPA.Browser.Selenium.Get Text
    ...    //*[@id="maincontent"]/div[6]/div[1]/div[1]/div/table/tbody/tr[3]/td[2]
    ${LastQuaterEarning}=    RPA.Browser.Selenium.Get Text
    ...    //*[@id="maincontent"]/div[6]/div[1]/div[1]/div/table/tbody/tr[5]/td[2]
    ${Year-agoEarning}=    RPA.Browser.Selenium.Get Text
    ...    //*[@id="maincontent"]/div[6]/div[1]/div[1]/div/table/tbody/tr[6]/td[2]
    ${Average Target Price}=    RPA.Browser.Selenium.Get Text
    ...    //*[@id="maincontent"]/div[6]/div[1]/div[1]/div/table/tbody/tr[2]/td[2]
    ${Current Year's Estimate}=    RPA.Browser.Selenium.Get Text
    ...    //*[@id="maincontent"]/div[6]/div[1]/div[1]/div/table/tbody/tr[8]/td[2]
    ${Median PE on CY Estimate}=    RPA.Browser.Selenium.Get Text
    ...    //*[@id="maincontent"]/div[6]/div[1]/div[1]/div/table/tbody/tr[9]/td[2]
    ${listexcel}=    Create List
    Append To List    ${listexcel}
    ...    ${companyname}
    ...    ${value}
    ...    ${ChangePoint}
    ...    ${percent}
    ...    ${finalresult}
    ...    ${Element color Final}
    ...    ${CloseValue}
    ...    ${changeValue}
    ...    ${percent1}
    ...    ${finalresult1}
    ...    ${Element color Final 1}
    ...    ${OPEN}
    ...    ${DAY RANGE}
    ...    ${52 WEEK RANGE}
    ...    ${MARKET CAP}
    ...    ${ SHARES OUTSTANDING}
    ...    ${PUBLIC FLOAT}
    ...    ${BETA}
    ...    ${REV. PER EMPLOYEE}
    ...    ${P/E RATIO}
    ...    ${EPS}
    ...    ${YIELD}
    ...    ${DIVIDEND}
    ...    ${EX-DIVIDEND DATE}
    ...    ${SHORT INTEREST}
    ...    ${% OF FLOAT SHORTED}
    ...    ${AVERAGE VOLUME}
    ...    ${Rating}
    ...    ${LastQuaterEarning}
    ...    ${Year-agoEarning}
    ...    ${Average Target Price}
    ...    ${Current Year's Estimate}
    ...    ${Median PE on CY Estimate}
    RETURN    ${listexcel}    ${ChangePoint}

save data into excel
    [Arguments]    ${item}    ${listexcel}    ${c}
    Open Workbook    AA500_Input.xlsm
    Read Worksheet    Sheet1
    Set Cell Value    ${c}    F    ${listexcel}[0]
    Set Cell Value    ${c}    G    ${listexcel}[1]
    Set Cell Value    ${c}    H    ${listexcel}[2]
    Set Cell Value    ${c}    I    ${listexcel}[3]
    Set Cell Value    ${c}    J    ${listexcel}[4]
    Set Cell Value    ${c}    K    ${listexcel}[5]
    Set Cell Value    ${c}    L    ${listexcel}[6]
    Set Cell Value    ${c}    M    ${listexcel}[7]
    Set Cell Value    ${c}    N    ${listexcel}[8]
    Set Cell Value    ${c}    O    ${listexcel}[9]
    Set Cell Value    ${c}    P    ${listexcel}[10]
    Set Cell Value    ${c}    Q    ${listexcel}[11]
    Set Cell Value    ${c}    R    ${listexcel}[12]
    Set Cell Value    ${c}    S    ${listexcel}[13]
    Set Cell Value    ${c}    T    ${listexcel}[14]
    Set Cell Value    ${c}    U    ${listexcel}[15]
    Set Cell Value    ${c}    V    ${listexcel}[16]
    Set Cell Value    ${c}    W    ${listexcel}[17]
    Set Cell Value    ${c}    X    ${listexcel}[18]
    Set Cell Value    ${c}    Y    ${listexcel}[19]
    Set Cell Value    ${c}    Z    ${listexcel}[20]
    Set Cell Value    ${c}    AA    ${listexcel}[21]
    Set Cell Value    ${c}    AB    ${listexcel}[22]
    Set Cell Value    ${c}    AC    ${listexcel}[23]
    Set Cell Value    ${c}    AD    ${listexcel}[24]
    Set Cell Value    ${c}    AE    ${listexcel}[25]
    Set Cell Value    ${c}    AF    ${listexcel}[26]
    Set Cell Value    ${c}    AG    ${listexcel}[27]
    Set Cell Value    ${c}    AH    ${listexcel}[28]
    Set Cell Value    ${c}    AI    ${listexcel}[29]
    Set Cell Value    ${c}    AJ    ${listexcel}[30]
    Set Cell Value    ${c}    AK    ${listexcel}[31]
    Set Cell Value    ${c}    AL    ${listexcel}[32]
    Save Workbook

Add data to Excel If no Change value present
    [Arguments]    ${listexcel 1}    ${c}    ${Exception}    ${item}
    Open Workbook    AA500_Input.xlsm
    Read Worksheet    Sheet1
    Set Cell Value    ${c}    F    ${listexcel 1}[0]
    Set Cell Value    ${c}    G    ${listexcel 1}[1]
    Set Cell Value    ${c}    H    ${listexcel 1}[2]
    Set Cell Value    ${c}    I    ${listexcel 1}[3]
    Set Cell Value    ${c}    J    ${listexcel 1}[4]
    Set Cell Value    ${c}    K    ${listexcel 1}[5]
    Set Cell Value    ${c}    L    ${Exception}
    Set Cell Value    ${c}    M    ${Exception}
    Set Cell Value    ${c}    N    ${Exception}
    Set Cell Value    ${c}    O    ${Exception}
    Set Cell Value    ${c}    P    ${Exception}
    Set Cell Value    ${c}    Q    ${listexcel 1}[6]
    Set Cell Value    ${c}    R    ${listexcel 1}[7]
    Set Cell Value    ${c}    S    ${listexcel 1}[8]
    Set Cell Value    ${c}    T    ${listexcel 1}[9]
    Set Cell Value    ${c}    U    ${listexcel 1}[10]
    Set Cell Value    ${c}    V    ${listexcel 1}[11]
    Set Cell Value    ${c}    W    ${listexcel 1}[12]
    Set Cell Value    ${c}    X    ${listexcel 1}[13]
    Set Cell Value    ${c}    Y    ${listexcel 1}[14]
    Set Cell Value    ${c}    Z    ${listexcel 1}[15]
    Set Cell Value    ${c}    AA    ${listexcel 1}[16]
    Set Cell Value    ${c}    AB    ${listexcel 1}[17]
    Set Cell Value    ${c}    AC    ${listexcel 1}[18]
    Set Cell Value    ${c}    AD    ${listexcel 1}[19]
    Set Cell Value    ${c}    AE    ${listexcel 1}[20]
    Set Cell Value    ${c}    AF    ${listexcel 1}[21]
    Set Cell Value    ${c}    AG    ${listexcel 1}[22]
    Set Cell Value    ${c}    AH    ${listexcel 1}[23]
    Set Cell Value    ${c}    AI    ${listexcel 1}[24]
    Set Cell Value    ${c}    AJ    ${listexcel 1}[25]
    Set Cell Value    ${c}    AK    ${listexcel 1}[26]
    Set Cell Value    ${c}    AL    ${listexcel 1}[27]
    Save Workbook

Save Exception in Excel When wrong company keyword enter
    [Arguments]    ${item}    ${Exception}    ${c}    ${nocompany}
    Open Workbook    AA500_Input.xlsm
    Read Worksheet    Sheet1
    Set Cell Value    ${c}    F    ${nocompany}
    Set Cell Value    ${c}    G    ${Exception}
    Set Cell Value    ${c}    H    ${Exception}
    Set Cell Value    ${c}    I    ${Exception}
    Set Cell Value    ${c}    J    ${Exception}
    Set Cell Value    ${c}    K    ${Exception}
    Set Cell Value    ${c}    L    ${Exception}
    Set Cell Value    ${c}    M    ${Exception}
    Set Cell Value    ${c}    N    ${Exception}
    Set Cell Value    ${c}    O    ${Exception}
    Set Cell Value    ${c}    P    ${Exception}
    Set Cell Value    ${c}    Q    ${Exception}
    Set Cell Value    ${c}    R    ${Exception}
    Set Cell Value    ${c}    S    ${Exception}
    Set Cell Value    ${c}    T    ${Exception}
    Set Cell Value    ${c}    U    ${Exception}
    Set Cell Value    ${c}    V    ${Exception}
    Set Cell Value    ${c}    W    ${Exception}
    Set Cell Value    ${c}    X    ${Exception}
    Set Cell Value    ${c}    Y    ${Exception}
    Set Cell Value    ${c}    Z    ${Exception}
    Set Cell Value    ${c}    AA    ${Exception}
    Set Cell Value    ${c}    AB    ${Exception}
    Set Cell Value    ${c}    AC    ${Exception}
    Set Cell Value    ${c}    AD    ${Exception}
    Set Cell Value    ${c}    AE    ${Exception}
    Set Cell Value    ${c}    AF    ${Exception}
    Set Cell Value    ${c}    AG    ${Exception}
    Set Cell Value    ${c}    AH    ${Exception}
    Set Cell Value    ${c}    AI    ${Exception}
    Set Cell Value    ${c}    AJ    ${Exception}
    Set Cell Value    ${c}    AK    ${Exception}
    Set Cell Value    ${c}    AL    ${Exception}
    Log    ${c}
    Save Workbook

Click Home Page
    Wait Until Keyword Succeeds    20x    0.5s    Click Element    //*[@id="Layer_1"]

Extract the data IF no analyst tab and change value Present
    [Arguments]    ${item}
    Wait Until Keyword Succeeds    40x    0.5s    Click Element    css:[class="primary j-result-link"]
    ${companyname}=    RPA.Browser.Selenium.Get Text    css:[class='company__name']
    ${value}=    RPA.Browser.Selenium.Get Text    css:[class='value']
    ${ChangePoint}=    RPA.Browser.Selenium.Get Text    css:[class='change--point--q']
    Log    ${ChangePoint}
    ${changepercent}=    RPA.Browser.Selenium.Get Text    css:[class='change--percent--q']
    Log    ${changepercent}
    ${percent}=    Remove String    ${changepercent}    %
    Log    ${percent}
    Sleep    5s
    Wait Until Keyword Succeeds    50x    0.5s    RPA.Desktop.Press Keys    Space
    Sleep    5s
    open application    calc.exe
    Wait Until Keyword Succeeds    40x    0.5s    RPA.Windows.Click    id:clearEntryButton
    Type Text    id:num${ChangePoint}Button
    RPA.Windows.Click    id:plusButton
    Type Text    id:num${percent}Button
    RPA.Windows.Click    id:equalButton
    ${result}=    Get Attribute    id:CalculatorResults    Name
    Log    ${result}
    Remove String    ${result}    Display is    null
    Log    ${result}
    Sleep    3s
    RPA.Windows.Click    id:Close
    ${finalresult}=    Replace String    ${result}    Display is    ${empty}
    Log    ${finalresult}
    ${Element color}=    Get Element Attribute    //*[@id="maincontent"]/div[2]/div[3]/div/div[2]/bg-quote    Class
    Log    ${Element color}
    ${Element color Final}=    Replace String    ${Element color}    intraday__change    ${empty}
    Wait Until Keyword Succeeds    40x    0.5s    RPA.Desktop.Press Keys    Up
    Wait Until Keyword Succeeds    40x    0.5s    RPA.Desktop.Press Keys    Up
    #Wait Until Keyword Succeeds    40x    0.5s    Click Element    //*[@id="maincontent"]/div[2]/div[4]/mw-chart/label
    ${ss}=    RPA.Browser.Selenium.Screenshot    //*[@id="maincontent"]/div[2]/div[4]    ${sspath}${item}.jpg
    ${OPEN}=    RPA.Browser.Selenium.Get Text
    ...    css:#maincontent > div.region.region--primary > div:nth-child(2) > div.group.group--elements.left > div.element.element--list > ul > li:nth-child(1) > span.primary
    ${DAY RANGE}=    RPA.Browser.Selenium.Get Text
    ...    css:#maincontent > div.region.region--primary > div:nth-child(2) > div.group.group--elements.left > div.element.element--list > ul > li:nth-child(2) > span.primary
    ${52 WEEK RANGE}=    RPA.Browser.Selenium.Get Text
    ...    css:#maincontent > div.region.region--primary > div:nth-child(2) > div.group.group--elements.left > div.element.element--list > ul > li:nth-child(3) > span.primary
    ${No_analyst_extractedlist}=    Create List
    Append To List
    ...    ${No_analyst_extractedlist}
    ...    ${companyname}
    ...    ${value}
    ...    ${ChangePoint}
    ...    ${percent}
    ...    ${finalresult}
    ...    ${Element color Final}
    ...    ${OPEN}
    ...    ${DAY RANGE}
    ...    ${52 WEEK RANGE}
    RETURN    ${No_analyst_extractedlist}    ${ChangePoint}

Save excel If no analyst tab and change value table
    [Arguments]    ${item}    ${No_analyst_extractedlist}    ${c}    ${Exception}
    Open Workbook    AA500_Input.xlsm
    Read Worksheet    Sheet1
    Set Cell Value    ${c}    F    ${No_analyst_extractedlist}[0]
    Set Cell Value    ${c}    G    ${No_analyst_extractedlist}[1]
    Set Cell Value    ${c}    H    ${No_analyst_extractedlist}[2]
    Set Cell Value    ${c}    I    ${No_analyst_extractedlist}[3]
    Set Cell Value    ${c}    J    ${No_analyst_extractedlist}[4]
    Set Cell Value    ${c}    K    ${No_analyst_extractedlist}[5]
    Set Cell Value    ${c}    L    ${Exception}
    Set Cell Value    ${c}    M    ${Exception}
    Set Cell Value    ${c}    N    ${Exception}
    Set Cell Value    ${c}    O    ${Exception}
    Set Cell Value    ${c}    P    ${Exception}
    Set Cell Value    ${c}    Q    ${No_analyst_extractedlist}[6]
    Set Cell Value    ${c}    R    ${No_analyst_extractedlist}[7]
    Set Cell Value    ${c}    S    ${No_analyst_extractedlist}[8]
    Set Cell Value    ${c}    T    ${Exception}
    Set Cell Value    ${c}    U    ${Exception}
    Set Cell Value    ${c}    V    ${Exception}
    Set Cell Value    ${c}    W    ${Exception}
    Set Cell Value    ${c}    X    ${Exception}
    Set Cell Value    ${c}    Y    ${Exception}
    Set Cell Value    ${c}    Z    ${Exception}
    Set Cell Value    ${c}    AA    ${Exception}
    Set Cell Value    ${c}    AB    ${Exception}
    Set Cell Value    ${c}    AC    ${Exception}
    Set Cell Value    ${c}    AD    ${Exception}
    Set Cell Value    ${c}    AE    ${Exception}
    Set Cell Value    ${c}    AF    ${Exception}
    Set Cell Value    ${c}    AG    ${Exception}
    Set Cell Value    ${c}    AH    ${Exception}
    Set Cell Value    ${c}    AI    ${Exception}
    Set Cell Value    ${c}    AJ    ${Exception}
    Set Cell Value    ${c}    AK    ${Exception}
    Set Cell Value    ${c}    AL    ${Exception}
    Log    ${c}
    Save Workbook

Click If Search Value Present
    ${companynamepresent}=    Does Page Contain Element    css:[class="primary j-result-link"]
    Log    ${companynamepresent}
    RETURN    ${companynamepresent}

Data Extraction If Anylist Tab Present and Change Value not Present
    [Arguments]    ${item}
    ${companyname}=    RPA.Browser.Selenium.Get Text    css:[class='company__name']
    ${value}=    RPA.Browser.Selenium.Get Text    css:[class='value']
    ${ChangePoint}=    RPA.Browser.Selenium.Get Text    css:[class='change--point--q']
    Log    ${ChangePoint}
    ${changepercent}=    RPA.Browser.Selenium.Get Text    css:[class='change--percent--q']
    Log    ${changepercent}
    ${percent}=    Remove String    ${changepercent}    %
    Log    ${percent}
    Sleep    5s
    Wait Until Keyword Succeeds    40x    0.5s    RPA.Desktop.Press Keys    Space
    Sleep    5s
    open application    calc.exe
    Wait Until Keyword Succeeds    40x    0.5s    RPA.Windows.Click    id:clearEntryButton
    Type Text    id:num${ChangePoint}Button
    RPA.Windows.Click    id:plusButton
    Type Text    id:num${percent}Button
    RPA.Windows.Click    id:equalButton
    ${result}=    Get Attribute    id:CalculatorResults    Name
    Log    ${result}
    Remove String    ${result}    Display is    null
    Log    ${result}
    Sleep    3s
    RPA.Windows.Click    id:Close
    ${finalresult}=    Replace String    ${result}    Display is    ${empty}
    Log    ${finalresult}
    ${Element color}=    Get Element Attribute    //*[@id="maincontent"]/div[2]/div[3]/div/div[2]/bg-quote    Class
    Log    ${Element color}
    ${Element color Final}=    Replace String    ${Element color}    intraday__change    ${empty}
    Wait Until Keyword Succeeds    40x    0.5s    RPA.Desktop.Press Keys    Up
    Wait Until Keyword Succeeds    40x    0.5s    RPA.Desktop.Press Keys    Up
    ${ss}=    RPA.Browser.Selenium.Screenshot    //*[@id="maincontent"]/div[2]/div[4]    ${sspath}${item}.jpg
    ${OPEN}=    RPA.Browser.Selenium.Get Text    //*[@id="maincontent"]/div[6]/div[1]/div[1]/div/ul/li[1]/span[1]
    ${DAY RANGE}=    RPA.Browser.Selenium.Get Text
    ...    css:#maincontent > div.region.region--primary > div:nth-child(2) > div.group.group--elements.left > div > ul > li:nth-child(2) > span.primary
    ${52 WEEK RANGE}=    RPA.Browser.Selenium.Get Text
    ...    css:#maincontent > div.region.region--primary > div:nth-child(2) > div.group.group--elements.left > div > ul > li:nth-child(3) > span.primary
    ${MARKET CAP}=    RPA.Browser.Selenium.Get Text
    ...    css:#maincontent > div.region.region--primary > div:nth-child(2) > div.group.group--elements.left > div > ul > li:nth-child(4) > span.primary
    ${ SHARES OUTSTANDING}=    RPA.Browser.Selenium.Get Text
    ...    css:#maincontent > div.region.region--primary > div:nth-child(2) > div.group.group--elements.left > div > ul > li:nth-child(5) > span.primary
    ${PUBLIC FLOAT}=    RPA.Browser.Selenium.Get Text
    ...    css:#maincontent > div.region.region--primary > div:nth-child(2) > div.group.group--elements.left > div > ul > li:nth-child(6) > span.primary
    ${BETA}=    RPA.Browser.Selenium.Get Text
    ...    css:#maincontent > div.region.region--primary > div:nth-child(2) > div.group.group--elements.left > div > ul > li:nth-child(7) > span.primary
    ${REV. PER EMPLOYEE}=    RPA.Browser.Selenium.Get Text
    ...    css:#maincontent > div.region.region--primary > div:nth-child(2) > div.group.group--elements.left > div > ul > li:nth-child(8) > span.primary
    ${P/E RATIO}=    RPA.Browser.Selenium.Get Text
    ...    css:#maincontent > div.region.region--primary > div:nth-child(2) > div.group.group--elements.left > div > ul > li:nth-child(9) > span.primary
    ${EPS}=    RPA.Browser.Selenium.Get Text
    ...    css:#maincontent > div.region.region--primary > div:nth-child(2) > div.group.group--elements.left > div > ul > li:nth-child(10) > span.primary
    ${YIELD}=    RPA.Browser.Selenium.Get Text
    ...    css:#maincontent > div.region.region--primary > div:nth-child(2) > div.group.group--elements.left > div > ul > li:nth-child(11) > span.primary
    ${DIVIDEND}=    RPA.Browser.Selenium.Get Text
    ...    css:#maincontent > div.region.region--primary > div:nth-child(2) > div.group.group--elements.left > div > ul > li:nth-child(12) > span.primary
    ${EX-DIVIDEND DATE}=    RPA.Browser.Selenium.Get Text
    ...    css:#maincontent > div.region.region--primary > div:nth-child(2) > div.group.group--elements.left > div > ul > li:nth-child(13) > span.primary
    ${SHORT INTEREST}=    RPA.Browser.Selenium.Get Text
    ...    css:#maincontent > div.region.region--primary > div:nth-child(2) > div.group.group--elements.left > div > ul > li:nth-child(14) > span.primary
    ${% OF FLOAT SHORTED}=    RPA.Browser.Selenium.Get Text
    ...    css:#maincontent > div.region.region--primary > div:nth-child(2) > div.group.group--elements.left > div > ul > li:nth-child(15) > span.primary
    ${AVERAGE VOLUME}=    RPA.Browser.Selenium.Get Text
    ...    css:#maincontent > div.region.region--primary > div:nth-child(2) > div.group.group--elements.left > div > ul > li:nth-child(16) > span.primary
    Wait Until Keyword Succeeds
    ...    20x
    ...    0.5s
    ...    Click Element
    ...    //*[@id="maincontent"]/div[5]/div/div/li[6]/a
    ${Rating}=    RPA.Browser.Selenium.Get Text
    ...    //*[@id="maincontent"]/div[6]/div[1]/div[1]/div/table/tbody/tr[3]/td[2]
    ${LastQuaterEarning}=    RPA.Browser.Selenium.Get Text
    ...    //*[@id="maincontent"]/div[6]/div[1]/div[1]/div/table/tbody/tr[5]/td[2]
    ${Year-agoEarning}=    RPA.Browser.Selenium.Get Text
    ...    //*[@id="maincontent"]/div[6]/div[1]/div[1]/div/table/tbody/tr[6]/td[2]
    ${Average Target Price}=    RPA.Browser.Selenium.Get Text
    ...    //*[@id="maincontent"]/div[6]/div[1]/div[1]/div/table/tbody/tr[2]/td[2]
    ${Current Year's Estimate}=    RPA.Browser.Selenium.Get Text
    ...    //*[@id="maincontent"]/div[6]/div[1]/div[1]/div/table/tbody/tr[8]/td[2]
    ${Median PE on CY Estimate}=    RPA.Browser.Selenium.Get Text
    ...    //*[@id="maincontent"]/div[6]/div[1]/div[1]/div/table/tbody/tr[9]/td[2]
    #Wait Until Keyword Succeeds    40x    0.5s    Click Element    //*[@id="maincontent"]/div[2]/div[4]/mw-chart/label
    ${listexcel 1}=    Create List
    Append To List    ${listexcel 1}
    ...    ${companyname}
    ...    ${value}
    ...    ${ChangePoint}
    ...    ${percent}
    ...    ${finalresult}
    ...    ${Element color Final}
    ...    ${OPEN}
    ...    ${DAY RANGE}
    ...    ${52 WEEK RANGE}
    ...    ${MARKET CAP}
    ...    ${ SHARES OUTSTANDING}
    ...    ${PUBLIC FLOAT}
    ...    ${BETA}
    ...    ${REV. PER EMPLOYEE}
    ...    ${P/E RATIO}
    ...    ${EPS}
    ...    ${YIELD}
    ...    ${DIVIDEND}
    ...    ${EX-DIVIDEND DATE}
    ...    ${SHORT INTEREST}
    ...    ${% OF FLOAT SHORTED}
    ...    ${AVERAGE VOLUME}
    ...    ${Rating}
    ...    ${LastQuaterEarning}
    ...    ${Year-agoEarning}
    ...    ${Average Target Price}
    ...    ${Current Year's Estimate}
    ...    ${Median PE on CY Estimate}
    RETURN    ${listexcel 1}    ${ChangePoint}

Check Change Value Present or not
    ${changevalue}=    Does Page Contain Element
    ...    //*[@id="maincontent"]/div[2]/div[3]/div/div[4]/table/tbody/tr/td[2]
    RETURN    ${changevalue}

Paste Graph in excel
    [Arguments]    ${item}    ${changepoint}
    Open Workbook    AA500_Input.xlsm
    IF    ${changepoint} > 0
        ${worksheet exists}=    Worksheet Exists    ${item}
        IF    ${worksheet exists} == True
            Remove Worksheet    ${item}
            Create Worksheet    ${item}
            Insert Image To Worksheet    5    E    ${sspath}${item}.jpg
            Save Workbook
        ELSE
            Create Worksheet    ${item}
            Insert Image To Worksheet    5    E    ${sspath}${item}.jpg
            Save Workbook
        END
    ELSE
        ${worksheet exists}=    Worksheet Exists    ${item}
        IF    ${worksheet exists} == True
            Remove Worksheet    ${item}
            Save Workbook
        ELSE
            Log    No ${item} worksheet Present
        END
    END
