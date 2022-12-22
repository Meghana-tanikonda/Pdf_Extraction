*** Settings ***
Library           RPA.Database
Library           Collections
Library           RPA.Browser.Selenium
#Library          RPA.Robocloud.Items
Library           OperatingSystem
Library           RPA.Robocorp.WorkItems
Library           RPA.Outlook.Application
Library           RPA.Archive

*** Variables ***
${hf}=            EMAIL ADDRES
${present}=       column is present
${not_present}=    column is not_present
${All_files}      C:/Users/meghana.tanikonda/Downloads/Samples.Zip
${fol1}=          C:${/}Users${/}meghana.tanikonda${/}Downloads${/}Unzippedfiles${/}Digital
${fol2}=          C:${/}Users${/}meghana.tanikonda${/}Downloads${/}Unzippedfiles${/}Scanned

*** Tasks ***
main
    # Checking Column names
    Download the outlook attachment
    ${list}=    Unzip a folder
    Creating List    ${list}

*** Keywords ***
Download the outlook attachment
    Open Application    ${True}
    ${emails}=    Get Emails
    ...    meghana.tanikonda@yash.com
    ...    Inbox
    ...    [Subject]='PDF_Operation'
    ...    ${True}
    ...    C:${/}Users${/}meghana.tanikonda${/}Downloads
    ...
    ...    ${True}
    ...    Received

Unzip a folder
    Extract Archive    ${All_files}    C:${/}Users${/}meghana.tanikonda${/}Downloads${/}Unzippedfiles
    ${List}=    Create List
    ...    C:${/}Users${/}meghana.tanikonda${/}Downloads${/}Unzippedfiles${/}Digital
    ...    C:${/}Users${/}meghana.tanikonda${/}Downloads${/}Unzippedfiles${/}Scanned
    RETURN    ${List}

Creating List
    [Arguments]    ${list}
    #${list}=    Create List    ${fol1}    ${fol2}
    ${counter}=    Set Variable    0
    ${items}=    Create List
    FOR    ${i}    IN    @{list}
        ${counter}=    Evaluate    ${counter}+1
        ${paths}=    List Files In Directory    ${i}
        FOR    ${file}    IN    @{paths}
            Append To List    ${items}    ${i}/${file}
        END
    END
    FOR    ${j}    IN    @{items}
        ${Filep}=    Set Variable    ${j}
        ${dict}=    Create Dictionary    name=${Filep}
        Create Output Work Item    ${dict}    save=True
    END
