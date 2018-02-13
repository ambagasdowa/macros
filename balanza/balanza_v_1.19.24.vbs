Private Sub ComboBox1_Change()

End Sub

Private Sub CommandButton1_Click()

'--------------------------------------------------------------------------------------------------------------------------------------------------
' => @baizabal.jesus# ah ADD SET NOCOUNT ON; in your procedure before the AS Instruction
'--------------------------------------------------------------------------------------------------------------------------------------------------

    Dim SellStartDate As String  'Declare the SellStartDate as Date
    Dim SellEndDate As String    'Declare the SellEndDate as Date
    Dim Company As String 'Declare company
    Dim iniPer As Integer 'Declare the init period of the balance
    Dim endPer As Integer 'Declare the end Period of the Balance
    Dim curYear As String 'then the trick is in the definition
    Dim sheetName As String 'Set the name of the sheet where source is fetch
    Dim addStringQuery As String 'where to save the companies
    Dim delimiter As String 'how to know which companies you select for the query
    Dim Count As String
    Dim endDate As Integer

'--------------------------------------------------------------------------------------------------------------------------------------------------
' PROCEDURE SECTION - Adding procedure Support
'--------------------------------------------------------------------------------------------------------------------------------------------------

    Dim con As ADODB.Connection
    Dim cmd As ADODB.Command
    Dim rs As ADODB.Recordset
    Dim par As String
    Dim WSP As Worksheet

    Set con = New ADODB.Connection
    Set cmd = New ADODB.Command
    Set rs = New ADODB.Recordset

'--------------------------------------------------------------------------------------------------------------------------------------------------
' PROCEDURE SECTION - Adding procedure Support
'--------------------------------------------------------------------------------------------------------------------------------------------------

    'Initializing the status bar
    Application.DisplayStatusBar = True

    'The Delimiter is used in mssql as well
    delimiter = "|"
    sheetName = "source"
    'curYear = Format(Date, "yyyy") 'can use instead => year(date)
    curYear = Sheets(sheetName).Range("D5").Value
    curYear = curYear + 2
    selectedYear = Sheets(sheetName).Range("E" + curYear).Value

    iniPer = 1   'This is for companies
    endPer = 10  'Same hir
    endDate = 12 'The length of the period thi\'s Dec
    SellStartDate = Sheets(sheetName).Range("D4").Value   'Pass value from cell B3 to SellStartDate variable
    'SellEndDate = Sheets(sheetName).Range("D3").Value     'Pass value from cell B4 to SellEndDate variable
    SellEndDate = SellStartDate
    'Define the companies
    'Going to deeper in the object and extract the soup!

    'MsgBox curYear
    'MsgBox selectedYear
    'MsgBox ListBox1.ListCount
    'build the company string
    Count = 0

    For i = 0 To ListBox1.ListCount - 1
        If ListBox1.Selected(i) = True Then
            addStringQuery = addStringQuery + delimiter & Worksheets(sheetName).Range("B3").Offset(i, 0).Value
            'MsgBox ListBox1.Selected(i)
        Else
            If ListBox1.Selected(i) = False Then
                Count = Count + 1
            End If
        End If
    Next i

    'MsgBox addStringQuery
    'If No Selection of company made then
    'Exit of the sub and expect for new orders
    If Count = ListBox1.ListCount Then
        MsgBox "Seleccione al menos una Compania"
        Exit Sub
    End If

    Company = Right(addStringQuery, Len(addStringQuery) - 1)
    'MsgBox Company

    'see the containt of curYear
    'MsgBox curYear
    'Build the period array
    Dim perBuild(1 To 12)
    For i = iniPer To endDate
        perBuild(i) = selectedYear & Format(i, "00")
        'MsgBox perBuild(i)
    Next
    'MsgBox CompanyName(Company)
    'MsgBox Company
    'this is the query that we need to perform and have to be dinamic and this make the thinks going to be interesting
    'dbo.getBalanzaComprobacion('201503','201503','TBKORI|TBKRAM|TBKGDL|TBKLAP|TBKCUL|TBKHER') order by Cuenta asc;

'--------------------------------------------------------------------------------------------------------------------------------------------------
' PROCEDURE SECTION - add the super Procedure
'--------------------------------------------------------------------------------------------------------------------------------------------------
    Application.StatusBar = "Contacting SQL Server..."

    ' Remove any values in the cells where we want to put our Stored Procedure's results.
    Dim rngRange As Range
    Set rngRange = Range(Cells(9, 2), Cells(Rows.Count, 1)).EntireRow
    rngRange.ClearContents

    'MsgBox "Retrive the data set...", vbInformation

    'Insert a progress bar
    'Sheets("Balanza").Range("E20").Value = "Retrive the data set..."
    'rngRange.ClearContents

    ' Log into our SQL Server, and run the Stored Procedure
    'con.Open "Provider=SQLOLEDB;Data Source=localhost;Initial Catalog=Northwind;Integrated Security=SSPI;Trusted_Connection=Yes;"
    con.Open "Provider=SQLOLEDB;Password=secret;Persist Security Info=True;User ID=Integra;Initial Catalog=integraapp;Data Source=localhost;Use Procedure for Prepare=1;Auto Translate=True;Packet Size=4096;Workstation ID=MEGUIDDO;Use Encryption for Data=False;Tag with column collation when possible=False"

    cmd.ActiveConnection = con

    Dim prmCustomerID As ADODB.Parameter

    ' Set up the parameter for our Stored Procedure
    ' (Parameter types can be adVarChar,adDate,adInteger)
    'cmd.Parameters.Append cmd.CreateParameter("beginDate", adVarChar, adParamInput, 10, Range("D2").Text)
    cmd.Parameters.Append cmd.CreateParameter("beginDate", adVarChar, adParamInput, 10, perBuild(SellStartDate))
    cmd.Parameters.Append cmd.CreateParameter("endDate", adVarChar, adParamInput, 10, perBuild(SellEndDate))
    cmd.Parameters.Append cmd.CreateParameter("Company", adVarChar, adParamInput, 255, Trim(Company))
    cmd.Parameters.Append cmd.CreateParameter("Delimiter", adVarChar, adParamInput, 10, delimiter)

    Application.StatusBar = "Running stored procedure..."
    cmd.CommandText = "dbo.sp_udsp_getBalanzaComprobacion"
    Set rs = cmd.Execute(, , adCmdStoredProc)
    'MsgBox rs.EOF
    ' Copy the results to cell B7 on the first Worksheet
    Application.StatusBar = "Retrive the data set..."
    Set WSP = Worksheets(1)
    WSP.Activate
    If rs.EOF = False Then WSP.Cells(9, 2).CopyFromRecordset rs

    rs.Close
    Set rs = Nothing
    Set cmd = Nothing

    con.Close
    Set con = Nothing

    Application.StatusBar = "Data successfully updated."
'--------------------------------------------------------------------------------------------------------------------------------------------------
' PROCEDURE SECTION
'--------------------------------------------------------------------------------------------------------------------------------------------------


    'Pass the Parameters values to the Stored Function used in the MSSQLData Connection
        'With ActiveWorkbook.Connections("balanza").OLEDBConnection
        '.CommandText = "select Cuenta,ENTIDADES,empresa,Descripci�n,Inicial,Cargo,Cr�dito,Final from dbo.getBalanzaComprobacion ('" & perBuild(SellStartDate) & "','" & perBuild(SellEndDate) & "','" & Trim(Company) & "','" & delimiter & "') order by Cuenta asc"
        'ActiveWorkbook.Connections("balanza").Refresh
        'ActiveWorkbook.RefreshAll
        'End With
End Sub

'this is just for testing issues , and don't reallly use it , if you see this code is already in CommandButton1_Click()
Sub btnClick()
Dim addStringQuery As String
    For i = 0 To ListBox1.ListCount - 1
        If ListBox1.Selected(i) = True Then
            MsgBox "index => " & Worksheets("source").Range("B3").Offset(i, 0).Value
            addStringQuery = addStringQuery + "|" & Worksheets("source").Range("B3").Offset(i, 0).Value
        End If
    Next i
    addStringQuery = Right(addStringQuery, Len(addStringQuery) - 1)
    'MsgBox Right(addStringQuery, Len(addStringQuery) - 1)
    'MsgBox addStringQuery
    'System.Diagnostics.Debug.WriteLine (ListBox1.SelectedItem(1).Tostring())
End Sub


Private Sub ListBox1_Click()

End Sub
