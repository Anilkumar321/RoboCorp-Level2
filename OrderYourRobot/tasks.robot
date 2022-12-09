*** Settings ***
Documentation       Insert the sales data for the week and export it as a PDF.

Library             RPA.Browser.Selenium    auto_close=${FALSE}
Library             RPA.Excel.Files
Library             RPA.HTTP
Library             RPA.PDF
Library             RPA.Tables
Library             RPA.Archive


*** Tasks ***
Insert the sales data for the week and export it as a PDF
    Open the order website
    Download Orders List
    Fill the form using the data from the Excel file
    Close the Browser


*** Keywords ***
Open the order website
    Open Available Browser    https://robotsparebinindustries.com/#/robot-order
    Sleep    5s
    Click Button    OK
    Maximize Browser Window

Download Orders List
    Download    https://robotsparebinindustries.com/orders.csv    overwrite=True

Fill and submit the form for one person
    [Arguments]    ${sales_rep}
    TRY
        Select From List By Index    head    ${sales_rep}[Head]
        Select Radio Button    body    ${sales_rep}[Body]
        Input Text    xpath=//input[@placeholder="Enter the part number for the legs"]    ${sales_rep}[Legs]
        Input Text    address    ${sales_rep}[Address]
        Click Button    order
        Wait Until Element Is Visible    id:receipt
        Screenshot    id:receipt    Screenshots${/}${sales_rep}[Order number].png
        Wait Until Element Is Visible    id:receipt
        ${sales_results_html}=    Get Element Attribute    id:receipt    outerHTML
        Html To Pdf    ${sales_results_html}    Receipt PDFs${/}${sales_rep}[Order number].pdf
        Click Button    order-another
        ${screenshotimg}=    Create List    Screenshots${/}${sales_rep}[Order number].png
        Add Files To Pdf
        ...    ${screenshotimg}
        ...    Receipt PDFs${/}${sales_rep}[Order number].pdf    append=True
        Sleep    3s
        Click Button    OK
    EXCEPT
        Close Browser
        Open the order website
    END

Fill the form using the data from the Excel file
    ${sales_reps}=    Read table from CSV    orders.csv    header=True
    FOR    ${sales_rep}    IN    @{sales_reps}
        Fill and submit the form for one person    ${sales_rep}
    END

Close the Browser
    Close Browser
    Archive Folder With Zip
    ...    C:\\Users\\masheta\\OneDrive - Baxter\\Desktop\\OrderYourRobot\\Receipt PDFs
    ...    OrderReceipts.zip
