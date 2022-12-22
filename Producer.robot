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
${All_files}      C:/Users/meghana.tanikonda/Downloads/Samples.Zip

*** Tasks ***
main
    # Checking Column names
    Download the outlook attachment
    ${list}=    Unzip a folder
    Creating List    ${list}

*** Keywords ***
Download the outlook attachment
   TRY
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
    EXCEPT    message
        Log    unable to download outlook attachment
    END

Unzip a folder
    Extract Archive    ${All_files}    C:${/}Users${/}meghana.tanikonda${/}Downloads${/}Unzippedfiles
    ${Length}=    Get Length    ${All_files}
    Log    ${Length}
    IF    ${Length} > 0
        Log    ${Length}
        ${List}=    Create List
        ...    C:${/}Users${/}meghana.tanikonda${/}Downloads${/}Unzippedfiles${/}Digital
        ...    C:${/}Users${/}meghana.tanikonda${/}Downloads${/}Unzippedfiles${/}Scanned
    ELSE
        Log    No files found
    END
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
