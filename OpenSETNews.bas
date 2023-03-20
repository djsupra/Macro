Attribute VB_Name = "OpenSETNews"
Public Function Escape(ByVal param As String) As String

    Dim i As Integer, BadChars As String

    
    param = Replace(param, " ", "_")
    'param = Replace(param, ",", "_")
    param = Replace(param, "/", "-")
    param = Replace(param, "\", "-")
    param = Replace(param, ":", "-")
    param = Replace(param, "*", "-")
    
    'BadChars = "%<>=&!@#$^()+{[}]|\;:'"",/?"
    BadChars = "%<>=!@#$^+{[}]|;'""?"
    
    For i = 1 To Len(BadChars)
        param = Replace(param, Mid(BadChars, i, 1), "%" & Hex(Asc(Mid(BadChars, i, 1))))
    Next
    
    Escape = param

End Function

Sub OpenURL_Click()

    '==================================
    'Delete old connection in Workbook Excel
    
    Dim Conn As WorkbookConnection
    For Each Conn In ThisWorkbook.Connections
        If Left(Conn.Name, 1) = "9" Or _
            Left(Conn.Name, 8) = "newslist" Then
            Conn.Delete
        End If
    Next Conn

    '==================================
    
     With Sheets("ELCID_List")
            .Range("H22:J2000").ClearContents
            .Range("N22:S2000").ClearContents
            '.Range("A30:K120").Clear
     End With
     
     
    'Dim i As Long
    'Dim FileNum_X As Long
    'Dim FileData_X() As Byte
    Dim PageAddress_X As String
    Dim PageSource_X As String
    Dim PageSource_X1 As String
    Dim PageSource_X2 As String
    'Dim PDFAddress_X As String
    Dim PageHTTP_X As Object
     
    On Error Resume Next
        Set PageHTTP_X = CreateObject("WinHTTP.WinHTTPrequest.5")
        If Err.Number <> 0 Then
            Set PageHTTP_X = CreateObject("WinHTTP.WinHTTPrequest.5.1")
        End If
    On Error GoTo 0

    With Sheets("Configuration")

            'PageAddress_X = .Range("C1")
        PageAddress_X = "https://classic.set.or.th/set/newslist.do?headline=&to=31%2F08%2F2015&symbol=&submit=%E0%B8%84%E0%B9%89%E0%B8%99%E0%B8%AB%E0%B8%B2&source=&newsGroupId=&securityType=S&from=05%2F08%2F2015&language=th&currentpage=0&country=TH"     
	'PageHTTP_X.Open "GET", PageAddress_X, False
            'PageHTTP_X.Send
            'TestError
            'PageSource_X = PageHTTP_X.ResponseText
            
            'PageAddress_X = "https://classic.set.or.th/set/newslist.do?headline=&to=31%2F08%2F2015&symbol=&submit=%E0%B8%84%E0%B9%89%E0%B8%99%E0%B8%AB%E0%B8%B2&source=&newsGroupId=&securityType=S&from=05%2F08%2F2015&language=th&currentpage=0&country=TH"
            'PageAddress_X = "https://classic.set.or.th/set/newslist.do?source=&symbol=&securityType=S&newsGroupId=&headline="
                                            '%E0%B8%97%E0%B8%B5%E0%B9%88%E0%B8%9B%E0%B8
            'PageAddress_X1 = "%E0%B8%97%E0%B8%B5%E0%B9%88%E0%B8%9B"
            'PageAddress_X2 = "&from=05%2F08%2F2015&to=31%2F08%2F2015&submit=%E0%B8%84%E0%B9%89%E0%B8%99%E0%B8%AB%E0%B8%B2&language=th&country=TH#content"
            'PageAddress_X = PageAddress_X & PageAddress_X1 & PageAddress_X2
            
                
            'Set IE = CreateObject("InternetExplorer.Application")
            'IE.Visible = True
            'IE.Navigate2 PageAddress_X
            'Do While IE.Busy = True
            '        DoEvents
            'Loop
              
            
            'Sheets("Text_X").Visible = xlSheetHidden
            Sheets("Text_X").Range("A:A").ClearContents
            'MsgBox PageSource_X
            'Sheets("Paper_X").Range("A1") = PageSource_X
            
            'MsgBox PageAddress_X
            'PageAddress_X = Escape(PageAddress_X)
            'MsgBox PageAddress_X
            
            With Sheets("Text_X").QueryTables.Add(Connection:= _
                    "Text;" & PageAddress_X, Destination:=Sheets("Text_X").Range("A1"))
                            .Name = "ELCID_Connection"
                            .FieldNames = True
                            .RowNumbers = False
                            .FillAdjacentFormulas = False
                            .PreserveFormatting = False   'True
                            .RefreshOnFileOpen = False
                            .BackgroundQuery = True
                            .RefreshStyle = xlOverwriteCells
                            .SavePassword = False
                            .SaveData = True
                            .AdjustColumnWidth = False
                            .RefreshPeriod = 0
                            
                            .TextFilePromptOnRefresh = False
                            '.TextFilePlatform = 1252
                            .TextFilePlatform = 65001
                            .TextFileStartRow = 1
                            .TextFileParseType = xlDelimited
                            .TextFileTextQualifier = xlTextQualifierDoubleQuote
                            
                            .TextFileConsecutiveDelimiter = False   'True
                            .TextFileTabDelimiter = False   'True
                            .TextFileSemicolonDelimiter = False
                            .TextFileCommaDelimiter = False
                            .TextFileSpaceDelimiter = False
                            
                            .TextFileColumnDataTypes = Array(1, 1, 1) 'For each column you want to import you need a ",1"
                            .TextFileTrailingMinusNumbers = True
                            
                            .Refresh BackgroundQuery:=False
                            .Delete
            End With
            
            'IE.Quit
            
    End With

    Set PageHTTP_X = Nothing
    'MsgBox "Open the folder [ C:\Users\Supara\Downloads\ResearchPDF\ ] for the downloaded file..."
        
    
End Sub
