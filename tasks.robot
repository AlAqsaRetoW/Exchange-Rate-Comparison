*** Settings ***
Documentation       Get table and make excel

Library             RPA.Browser.Selenium
Library             RPA.Excel.Files
Library             DateTime
Library             html_tables.py

Task Teardown       Close All Browsers


*** Variables ***
${status}       Run Keyword and Return Status    Open Workbook    kurs_${date}.xlsx
${date}         Get Current Date    local    0    %d.%m.%Y    true


*** Tasks ***
Run Program
    Set Selenium Timeout    20
    ${date}=    Get Current Date    local    0    %d.%m.%Y    true
    Scrape Data and Make Excel BCA
    Scrape Data and Make Excel BRI
    Scrape Data and Make Excel Mandiri
    Process Excel


*** Keywords ***
Get HTML table BCA
    Open Available Browser
    ...    https://www.bca.co.id/id/informasi/kurs
    ...    headless=True
    ${html_table}=
    ...    Get Element Attribute
    ...    xpath:(//table[@class='m-table-kurs m-table--sticky-first-coloumn m-table-kurs--pad'])[1]
    ...    outerHTML
    RETURN    ${html_table}

Get HTML table BRIE
    Open Available Browser
    ...    https://bri.co.id/kurs-detail
    ...    headless=True
    Wait Until Element Is Visible    id:_bri_kurs_detail_portlet_display2
    ${html_table}=
    ...    Get Element Attribute
    ...    id:_bri_kurs_detail_portlet_display2
    ...    outerHTML
    RETURN    ${html_table}

Get HTML table BRITT
    Wait Until Element Is Visible    xpath:(//a[normalize-space()='KURS TT COUNTER'])[1]
    Click Element    xpath:(//a[normalize-space()='KURS TT COUNTER'])[1]
    ${html_table}=
    ...    Get Element Attribute
    ...    id:_bri_kurs_detail_portlet_display
    ...    outerHTML
    RETURN    ${html_table}

Get HTML table Mandiri
    Open Available Browser
    ...    https://www.bankmandiri.co.id/kurs
    ...    headless=True
    Wait Until Element Is Visible
    ...    id:_Exchange_Rate_Portlet_INSTANCE_9070nSEKk62r_display
    ${html_table}=
    ...    Get Element Attribute
    ...    id:_Exchange_Rate_Portlet_INSTANCE_9070nSEKk62r_display
    ...    outerHTML
    RETURN    ${html_table}

Scrape Data and Make Excel BCA
    ${html_table}=    Get HTML table BCA
    ${table}=    Read Table From Html    ${html_table}
    ${date}=    Get Current Date    local    0    %d.%m.%Y    true
    Create Workbook    kurs_${date}.xlsx    sheet_name=BCA
    FOR    ${row}    IN    @{table}
        Append rows to worksheet    ${row}
    END

Scrape Data and Make Excel BRI
    ${html_table}=    Get HTML table BRIE
    ${table}=    Read Table From Html    ${html_table}
    Create Worksheet    BRI
    FOR    ${row}    IN    @{table}
        Append rows to worksheet    ${row}
    END

    ${html_table}=    Get HTML table BRITT
    ${table}=    Read Table From Html    ${html_table}
    FOR    ${row}    IN    @{table}
        Append rows to worksheet    ${row}
    END

Scrape Data and Make Excel Mandiri
    ${html_table}=    Get HTML table Mandiri
    ${table}=    Read Table From Html    ${html_table}
    Create Worksheet    Mandiri
    FOR    ${row}    IN    @{table}
        Append rows to worksheet    ${row}
    END

Process Excel
    Create Worksheet    Compare

    Set Cell Value    1    2    Mata Uang
    Set Cell Value    1    3    USD

    Set Cell Value    3    3    Special Rate
    Set Cell Value    3    5    TT
    Set Cell Value    3    7    Bank Note

    Set Cell Value    4    3    Beli
    Set Cell Value    4    4    Jual
    Set Cell Value    4    5    Beli
    Set Cell Value    4    6    Jual
    Set Cell Value    4    7    Beli
    Set Cell Value    4    8    Jual
    Set Cell Value    4    2    Bank
    Set Cell Value    5    2    BCA
    Set Cell Value    6    2    BRI
    Set Cell Value    7    2    Mandiri

    # BCA eRate / special rate
    FOR    ${col_idx}    IN RANGE    2    4
        ${bca}=    Get Cell Value    3    ${col_idx}    BCA
        ${dest_col}=    Evaluate    ${col_idx} + 1
        Set Cell Value    5    ${dest_col}    ${bca}    Compare
    END
    # BCA TT
    FOR    ${col_idx}    IN RANGE    4    6
        ${bca}=    Get Cell Value    3    ${col_idx}    BCA
        ${dest_col}=    Evaluate    ${col_idx} + 1
        Set Cell Value    5    ${dest_col}    ${bca}    Compare
    END
    # BCA Bank Notes
    FOR    ${col_idx}    IN RANGE    6    8
        ${bca}=    Get Cell Value    3    ${col_idx}    BCA
        ${dest_col}=    Evaluate    ${col_idx} + 1
        Set Cell Value    5    ${dest_col}    ${bca}    Compare
    END

    # BRI eRate / special rate
    FOR    ${col_idx}    IN RANGE    2    4
        ${bri}=    Get Cell Value    2    ${col_idx}    BRI
        ${dest_col}=    Evaluate    ${col_idx} + 1
        Set Cell Value    6    ${dest_col}    ${bri}    Compare
    END
    # BRI TT
    FOR    ${col_idx}    IN RANGE    2    4
        ${bri}=    Get Cell Value    13    ${col_idx}    BRI
        ${dest_col}=    Evaluate    ${col_idx} + 3
        Set Cell Value    6    ${dest_col}    ${bri}    Compare
    END

    # Mandiri eRate / special rate
    FOR    ${col_idx}    IN RANGE    2    4
        ${mandiri}=    Get Cell Value    19    ${col_idx}    Mandiri
        ${dest_col}=    Evaluate    ${col_idx} + 1
        Set Cell Value    7    ${dest_col}    ${mandiri}    Compare
    END
    # Mandiri TT
    FOR    ${col_idx}    IN RANGE    4    6
        ${mandiri}=    Get Cell Value    19    ${col_idx}    Mandiri
        ${dest_col}=    Evaluate    ${col_idx} + 1
        Set Cell Value    7    ${dest_col}    ${mandiri}    Compare
    END
    # Mandiri Bank Notes
    FOR    ${col_idx}    IN RANGE    6    8
        ${mandiri}=    Get Cell Value    19    ${col_idx}    Mandiri
        ${dest_col}=    Evaluate    ${col_idx} + 1
        Set Cell Value    7    ${dest_col}    ${mandiri}    Compare
    END

    ${datetime}=    Get Current Date    local    0    timestamp    true

    Set Cell Value    9    2    Update
    Set Cell Value    9    3    ${datetime}

    Auto Size Columns    A    H
    Save Workbook
