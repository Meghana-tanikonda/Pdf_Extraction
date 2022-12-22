*** Settings ***
Documentation       Email and Pdf extraction.
...                 Download the outlook attachment.
...                 unzip the outlook attachment.
...                 A folder contains some pdf files and convert into text file and need to extract the data from text file
...                 store in Excel

Library             task.py
Library             RPA.Windows
Library             RPA.PDF
Library             RPA.Excel.Files
Library             String
Library             RPA.Outlook.Application    auto_close=${false}
Library             RPA.FileSystem
Library             RPA.RobotLogListener
Library             RPA.Archive
Library             RPA.Email.ImapSmtp
Library             RPA.Tables
Library             RPA.HTTP
Library             Collections
Library             RPA.Robocloud.Items


*** Variables ***
${sheetname}=       0


*** Tasks ***
Email and pdf extraction
    ${path1}=    For Each Input Work Item    load work items

    Create Workbook    C:/Users/meghana.tanikonda/Documents/Robotsparebin/workitem/output.xlsx
    FOR    ${i}    IN    @{path1}
        Log    ${i}

        ${data}=    Convert To String    ${i}
        ${pdf_data}=    Readpdf    ${data}
        ${len}=    Get Length    ${pdf_data}
        IF    ${len} == 1
            ${variable}=    Create Dictionary
            # ...    name=${file}
            ...    Value= Unable to extract scanned pdf data
            Create Worksheet    scanned
            Append Rows To Worksheet    ${variable}
            Rename worksheet    Sheet    Digital
            save Workbook    C:/Users/meghana.tanikonda/Documents/Robotsparebin/workitem/output.xlsx
        ELSE
            ${list_op}=    Extract the text data    ${pdf_data}
            store in excel    ${list_op}
        END
    END
    #Sending a mail


*** Keywords ***
Get the file name
    [Arguments]    ${path1}
    ${file_name}=    Get File Name    ${path1}
    ${data}=    Convert To String    ${path1}
    ${pdf_data}=    Readpdf    ${data}
    ${pdf_files}=    Get Text From Pdf    ${data}
    Log    ${pdf_data}
    RETURN    ${pdf_data}

Extract the text data
    [Arguments]    ${pdf_data}
    TRY
        ${Text_data}=    Convert To String    ${pdf_data}
        ${Date}=    Should Match Regexp    ${Text_data}    \\d{2}\\/\\d{2}\\/\\d{2}
        ${customer_Number}=    Should Match Regexp    ${Text_data}    C\\d+
        ${PoNumber}=    Should Match Regexp    ${Text_data}    \\d{10}
        ${InvoiceTotal}=    Should Match Regexp    ${Text_data}    \\$\\s+\\d+\\.\\d+
        ${Manager}=    Get Regexp Matches    ${Text_data}    (?sim)(?<=Account Manager: )\\w+\\s+\\w+
        ${order}=    Get Regexp Matches    ${Text_data}    (?sim)(?<=Order Taken By: )\\w+\\s+\\w+
        ${invoiceNumber}=    Should Match Regexp    ${Text_data}    \\w{2}\\s\\w{3}\\d{5}
        ${list_op1}=    Create Dictionary
        ...    Date=${Date}
        ...    customer_Number=${customer_Number}
        ...    PoNumber=${PoNumber}
        ...    invoicetotal=${InvoiceTotal}
        ...    Manager=${Manager}
        ...    order=${order}
        ...    invoiceNumber=${invoiceNumber}

        RETURN    ${list_op1}
    EXCEPT    message
        Log    unable to extract the data
    END

store in excel
    [Arguments]    ${final_tb}
    TRY
        set Worksheet Value    1    1    Date
        Set Worksheet Value    1    2    Custumer number
        Set Worksheet Value    1    3    PO number
        Set Worksheet Value    1    4    invoicetotal
        Set Worksheet Value    1    5    Manager
        Set Worksheet Value    1    6    order
        Set Worksheet Value    1    7    invoice number

        Append Rows To Worksheet    ${final_tb}

        save Workbook    C:/Users/meghana.tanikonda/Documents/Robotsparebin/workitem/output.xlsx
    EXCEPT    message
        Log    unable to store in excel
    END

Sending a mail
    TRY
        Send Email
        ...    recipients=meghana.tanikonda@yash.com
        ...    subject=Pdf extraction
        ...    body=Please find the updated excel
        ...    attachments=C:/Users/meghana.tanikonda/Documents/Robotsparebin/workitem/output.xlsx
    EXCEPT    message
        Log    unable to send a mail
    END

load work items
    ${work_items}=    Get Work Item Variables
    ${path1}=    Set Variable    ${work_items}[name]
    RETURN    ${path1}
