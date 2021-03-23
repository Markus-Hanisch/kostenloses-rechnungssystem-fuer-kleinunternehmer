Sub PrintRechnungAsPDF()
'
' PrintRechnungAsPDF Makro
  ' Simulierte Rechnung als PDF drucken

PathName = ActiveWorkbook.Path
RngNr = Tabelle5.Range("Rechnungsnummer").Value
RngNrLength = Len(RngNr)

If RngNrLength > 0 Then
    SvPdfAs = PathName & "\" & RngNr & ".pdf"
Else
    SvPdfAs = PathName & "\" & "Error" & ".pdf"
End If

Tabelle5.ExportAsFixedFormat Type:=xlTypePDF, Filename:=SvPdfAs, Quality:=xlQualityStandard, IncludeDocProperties:=False, IgnorePrintAreas:=False, OpenAfterPublish:=False

End Sub
