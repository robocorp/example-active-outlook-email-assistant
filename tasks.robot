*** Settings ***
Documentation       An Assistant Robot.

Library             OperatingSystem
Library             RPA.Assistant
Library             ExtendedOutlook    autoexit=${FALSE}


*** Tasks ***
Main
    Display Main Menu
    ${result}=    RPA.Assistant.Run Dialog
    ...    title=Assistant Demo
    ...    on_top=True
    ...    height=450


*** Keywords ***
Display Main Menu
    [Documentation]
    ...    Main UI of the bot. We use the "Back To Main Menu" keyword
    ...    with buttons to make other views return here.
    Clear Dialog
    Add Heading    Assistant Demo
    Add Text    Outlook:
    Add Button    Active Email    Show Active Email
    Add Submit Buttons    buttons=Close    default=Close

Back To Main Menu
    [Documentation]
    ...    This keyword handles the results of the form whenever the "Back" button
    ...    is used, and then we return to the main menu
    [Arguments]    ${results}={}

    # Handle the dialog results via the passed 'results' -variable
    # Logging the user outputs directly is bad practice as you can easily expose things that should not be exposed
    IF    'password' in ${results}    Log To Console    Do not log user inputs!
    IF    'files' in ${results}
        Log To Console    Selected files: ${results}[files]
    END

    Display Main Menu
    Refresh Dialog

Show Active Email
    Open Application
    ${email}=    Get Active Email
    Clear Dialog
    Add Next Ui Button    Back    Back To Main Menu
    IF    ${email}
        Add Heading    Active Email    Small
        ${email_text}=    Set Variable    SUBJECT: ${email}[Subject]\n
        ...    SENDER: ${email}[Sender]\n
        ...    RECEIVED: ${email}[ReceivedTime]\n
        ...    BODY\n----------------------\n${email}[Body][:80]\n----------------------
        Add Text    ${{''.join($email_text)}}
    ELSE
        Add Text    Nothing has been selected
    END
    Refresh Dialog
