Option Explicit 
Dim colArgs, outputFile, inputFile, inputName, inputFileList, fileName
Dim outputFullpath, inputFileFullpath, fileExists
Dim fileDelimiter, rawFile, fileExt, currentRow, rowText, rowColumnValue, columnValue, lineArray

' file system Object
Dim Objfso

' Create connection
Dim catDB, cnDB ' As ADODB.Connection
Dim objConn, strConnect
Dim ctable, sqlQuery

On Error Resume Next

' variable to hold arguments passed:
Set colArgs = WScript.Arguments.Named

' Arguments to be minimum 2.
If colArgs.Count < 2  Then
    Wscript.Echo " ******************************************************************** "
    WScript.Echo "Insufficient arguments supplied!."    
    Wscript.Echo "Required output (Access Database FileName). "
    Wscript.Echo "Required at-least one csv|pipe|tab file delimited. "
    Wscript.Echo "Usage: /o:<File Name with extension accdb> /i:<csv|pipe|tab FileName>,<csv|pipe|tab FileName>;...."
    Wscript.Echo "Example c:\temp>cscript.exe //nologo csv2access.vbs /o:output.accdb /i:s.csv,s2.pipe,s3.tab,s4.csv"
    Wscript.Echo " ******************************************************************** "
    WScript.Quit
End If

' Get the filename for the output excel
If colArgs.Exists("o") Then
    outputFile = colArgs.Item("o")
End If

' If the filename is blank or nothing
If (outputFile = Empty) Then
    Wscript.Echo "Usage: /o:<File Name with extension accdb> is required."
    Wscript.Quit
End If

' If the output extension is missing, then add xlsx as extension.
If (Instr(outputFile,".accdb")<=0) Then
    outputFile = outputFile & ".accdb"
End If

' Get the Input Files comma seperated
If colArgs.Exists("i") Then
    inputFile = colArgs.Item("i")
End If

' If the input filename is blank or nothing
If (inputFile = Empty) Then
    Wscript.Echo "Usage: /i:<csv|pipe|tab delimited filename> is required. "
    Wscript.Quit
End If

' create an Object to get the current Filesystem path.
Set Objfso = CreateObject("Scripting.FileSystemObject")
outputFullpath = Objfso.GetAbsolutePathName(outputFile)

fileExists=Objfso.FileExists(outputFullpath)
If fileExists Then
    Objfso.DeleteFile(outputFullpath) 
end If

strConnect = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source='" & outputFullpath & "'"
Set catDB = CreateObject("ADOX.Catalog")
catDB.Create strConnect

If err.Number = 0 Then
    wscript.echo "info: database created successfully"
   
    Set cnDB = catDB.ActiveConnection

Else
    wscript.echo "fail: error = " & err.Number & " (" & err.Description & ")"
    wscript.Quit(1)
End If

' process the input files
inputFileList = Split(inputFile,",")
For Each rawFile in inputFileList
    If (Instr(trim(rawFile),".csv") or Instr(Trim(rawFile),".tab") or Instr(Trim(rawFile),".pipe") ) Then
        inputFileFullpath = Objfso.GetAbsolutePathName(Trim(rawFile))
        fileExt = objfso.GetExtensionName(inputFileFullpath)
        fileName = replace(Objfso.GetFileName(inputFileFullpath),".csv","")
        fileName = replace(fileName,".tab","")
        fileName = replace(fileName,".pipe","")

        Set objConn = Nothing
        Set inputFile = Objfso.OpenTextFile(inputFileFullpath)
        currentRow = 1
        Do While inputFile.AtEndOfStream <> True
            If (UCase(fileExt) ="TAB") Then
                lineArray = Split(inputFile.ReadLine,vbTab)
            End If 
            If (UCase(fileExt) ="PIPE") Then
                lineArray = Split(inputFile.ReadLine,"|")
            End If
            If (UCase(fileExt) ="CSV") Then
                lineArray = Split(inputFile.ReadLine,",")
            End If
            For each columnValue in lineArray
                rowColumnValue = Trim(columnValue)
                'Check 1st position in column for doubleQuotes
                If (Left(rowColumnValue,1) = """") Then
                    rowColumnValue = Right(rowColumnValue,Len(rowColumnValue)-1)
                End If 
                'Check last column position for doubleQuotes
                If (right(rowColumnValue,1) = """") Then
                    rowColumnValue = Left(rowColumnValue,Len(rowColumnValue)-1)
                End If 
                If Trim(rowText) ="" Then
                    rowText = Trim(rowColumnValue)
                Else
                    rowText = rowText + "," + rowColumnValue
                End If
            Next
            lineArray = Split(rowText,",")
            'First Row check and create Table.
            If currentRow = 1 then
                For each columnValue in lineArray
                    ctable = ctable & ", " & columnValue &" TEXT(255) WITH COMPRESSION NULL"
                Next
                ctable  = "CREATE Table "& fileName & "( id int" & ctable &")"
                cnDB.Execute ctable
            End If
            If objConn is Nothing  Then
                Set objConn = CreateObject( "ADODB.Connection" )
                objConn.Open strConnect
            End If
            sqlQuery  = " " & (currentRow -1)
            If currentRow > 1 Then
                For each columnValue in lineArray
                    sqlQuery = sqlQuery & ",""" & columnValue & """"
                Next
                sqlQuery = "Insert into [" & fileName & "] values (" & sqlQuery & ")"
                objConn.Execute(sqlQuery)
            End If
            sqlQuery = Empty
            ctable = Empty
            rowText = Empty
            currentRow = currentRow + 1
        Loop
        inputFile.Close
    End If
Next

Set objConn = Nothing
Set cnDB = Nothing
Set catDB = Nothing
Set Objfso = Nothing

WScript.Echo "Done!"