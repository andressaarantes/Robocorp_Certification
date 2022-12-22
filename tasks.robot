*** Settings ***
Documentation       Orders robots from RobotSpareBin Industries Inc.
...                 Saves the order HTML receipt as a PDF file.
...                 Saves the screenshot of the ordered robot.
...                 Embeds the screenshot of the robot to the PDF receipt.
...                 Creates ZIP archive of the receipts and the images.

Library             RPA.Browser.Selenium    auto_close=${False}
Library             RPA.HTTP
Library             RPA.Excel.Files
Library             RPA.Tables
Library             RPA.Desktop
Library             RPA.PDF
Library             RPA.Archive
Library             RPA.Robocorp.Vault
Library             RPA.Dialogs
Library             Collections
Library             String


*** Variables ***
#${EXCEL_FILE_URL}=    https://robotsparebinindustries.com/orders.csv
${GLOBAL_RETRY_AMOUNT}      10x
${GLOBAL_RETRY_INTERVAL}    1s
${Receipt_print}
${files}
${zip_files}
${link_csv}


*** Tasks ***
Orders robots from RobotSpareBin Industries Inc
    Open the robot order website
    Fill the form using the csv data
    Creates ZIP archive of the receipts and the images


*** Keywords ***
Open the robot order website
    ${Website}=    Get Secret    Website_Order
    Open Headless Chrome Browser    ${Website}[site]
    Click Button    css:div > button.btn.btn-dark

Fill the form for one person
    [Arguments]    ${Order}
    Select From List By Value    head    ${Order}[Head]
    select Radio Button    body    ${Order}[Body]
    Input Text    class:form-control    ${Order}[Legs]
    Input Text    address    ${Order}[Address]

    Click button order and keep checking until success
    ${Receipt_print}=    Get Element Attribute    receipt    outerHTML
    Html To Pdf
    ...    ${Receipt_print}
    ...    ${OUTPUT_DIR}${/}receipts${/}${Order}[Order number].pdf
    Screenshot    robot-preview    ${OUTPUT_DIR}${/}receipts${/}${Order}[Order number].jpeg

    Open Pdf    ${OUTPUT_DIR}${/}receipts${/}${Order}[Order number].pdf
    ${files}=    Create List
    ...    ${OUTPUT_DIR}${/}receipts${/}${Order}[Order number].pdf
    ...    ${OUTPUT_DIR}${/}receipts${/}${Order}[Order number].jpeg
    Add Files To Pdf
    ...    ${files}
    ...    ${OUTPUT_DIR}${/}receipts${/}${Order}[Order number].pdf
    Close All Pdfs
    Click Button    order-another
    Click Button    css:div > button.btn.btn-dark

Fill the form using the csv data
    ${EXCEL_FILE_URL}=    Collect csv url from user
    Download    ${EXCEL_FILE_URL}    overwrite=True
    #Open File    orders.csv
    ${Orders}=    Read table from CSV    orders.csv    dialect=excel    header=True
    Log    Found columns: ${orders}
    # Open File    orders.csv
    FOR    ${Order}    IN    @{Orders}
        Log    ${Order}
        Fill the form for one person    ${Order}
    END
    #Close Workbook

Collect csv url from user
    Add text input    search    label=EXCEL_FILE_URL
    ${response}=    Run dialog
    RETURN    ${response.search}

Click button order and keep checking until success
    Wait Until Keyword Succeeds
    ...    ${GLOBAL_RETRY_AMOUNT}
    ...    ${GLOBAL_RETRY_INTERVAL}
    ...    Send Order

Send Order
    Click button    preview
    Click Button    order
    Page Should Contain Button    order-another

Creates ZIP archive of the receipts and the images
    Archive Folder With Zip    ${OUTPUT_DIR}${/}receipts    ${OUTPUT_DIR}${/}PDF receipts.zip    include=*.pdf
