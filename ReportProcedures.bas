Attribute VB_Name = "ReportProcedures"
Option Explicit

Public Sub FormatClientSheet(sht As Worksheet, hhName As String, eqTarget As String)
    With sht
        'Border at top of the page between content and header
        .Range("A1:F1").Borders(xlEdgeTop).Weight = xlMedium
        .Range("A1:F1").Borders(xlEdgeTop).LineStyle = xlContinuous
        .Cells.Font.Size = 11
        .Cells.Font.Name = "Arial"
        
        'Small space between header and content
        .Rows(1).RowHeight = 3
    End With
    
    'Format page layout
    AddHeader sht, hhName, eqTarget
    AddFooter sht
    FormatPrintArea sht
End Sub

Private Sub AddHeader(sht As Worksheet, hhName As String, eqTarget As String)
    With sht
        If hhName = "Schaefer, Russell & Patricia" Then
            'Client exception for the report
            .Range("A2").value = "S, Russell & Patricia"
        Else
            'Put in the household name(s)
            .Range("A2").value = hhName
        End If
        .Range("A2").Font.Bold = True
        
        If .Name = "Client Trades" Then
            .Range("A2:D2").Merge
            
            'Add the equity target
            .Range("E2").value = "Equity Target"
            .Range("E2").Font.Underline = True
            .Range("E3").value = eqTarget
            .Range("E3").HorizontalAlignment = xlLeft
        ElseIf .Name = "Trades by Subclass" Then
            .Range("A2:C2").Merge
            
            'Add the equity target
            .Range("D2").value = "Equity Target"
            .Range("D2").Font.Underline = True
            .Range("D3").value = eqTarget
            .Range("D3").HorizontalAlignment = xlLeft
        End If
            
        'Add the actual header
        With .PageSetup
            If Time < TimeValue("15:00:00") Then
                .LeftHeader = "&L" & chr(10) & chr(10) & "&""Arial""&12&BTrade Recommendations - " & Date
            Else
                If Weekday(Date, vbSunday) = 6 Then
                    .LeftHeader = "&L" & chr(10) & chr(10) & "&""Arial""&12&BTrade Recommendations - " & Date + 3
                Else
                    .LeftHeader = "&L" & chr(10) & chr(10) & "&""Arial""&12&BTrade Recommendations - " & Date + 1
                End If
            End If
        
            Dim imagePath As String
            imagePath = "Z:\DO NOT MOVE FPIS-logo-final.jpg"
            
            If Dir(imagePath) = "" Then
                ErrorHandling.ErrorAndContinue "FPIS logo not found. Header missing logo.", hhName:=hhName
            Else
                .RightHeader = "&R&g"
                .RightHeaderPicture.fileName = imagePath
                .RightHeaderPicture.LockAspectRatio = msoTrue
                .RightHeaderPicture.Height = Application.InchesToPoints(0.7)
            End If
        End With
    End With
End Sub

Private Sub AddFooter(sht As Worksheet)
    With sht.PageSetup
        Dim footerFormat As String
        Dim footerStr As String
        footerFormat = "&""Arial""&9&I"
        footerStr = "The recommendations outlined above are estimates subject to market fluctuations and may not be " _
            & "traded in the exact dollar value indicated."
        
        .LeftFooter = footerFormat & footerStr
    End With
End Sub

Private Sub FormatPrintArea(sht As Worksheet)
    'Turn off print communication for better performance.
    'This could make the header and footer faster as well, but they then don't fill properly for unknown reasons
    If Val(Application.Version) >= 14 And Left(Application.OperatingSystem, 3) = "Win" Then
        Application.PrintCommunication = False
    End If
    
    With sht.PageSetup
        .TopMargin = Application.InchesToPoints(1.2)
        .BottomMargin = Application.InchesToPoints(0.75)
        .LeftMargin = Application.InchesToPoints(0.4)
        .RightMargin = Application.InchesToPoints(0.4)
        .HeaderMargin = Application.InchesToPoints(0.5)
        .FooterMargin = Application.InchesToPoints(0.3)
    
        .Orientation = xlPortrait
    End With
    
    'Turn print communication back on
    If Val(Application.Version) >= 14 And Left(Application.OperatingSystem, 3) = "Win" Then
        Application.PrintCommunication = True
    End If
End Sub

Public Sub AddPageBreak(report As Worksheet, rng As Range)
    If Not rng Is Nothing Then
        'Add a page break before the account information
        report.Rows(rng.Offset(-1, 0).Row).PageBreak = xlPageBreakManual
        
        'Put a border at the top between the header and the content
        report.Range(rng, rng.Offset(-1, 5)).Borders(xlEdgeTop).LineStyle = xlContinuous
        report.Range(rng, rng.Offset(-1, 5)).Borders(xlEdgeTop).Weight = xlMedium
    End If
End Sub
