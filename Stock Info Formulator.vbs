{\rtf1\ansi\ansicpg1252\cocoartf2580
\cocoatextscaling0\cocoaplatform0{\fonttbl\f0\fswiss\fcharset0 Helvetica;}
{\colortbl;\red255\green255\blue255;}
{\*\expandedcolortbl;;}
\margl1440\margr1440\vieww11520\viewh8400\viewkind0
\pard\tx720\tx1440\tx2160\tx2880\tx3600\tx4320\tx5040\tx5760\tx6480\tx7200\tx7920\tx8640\pardirnatural\partightenfactor0

\f0\fs24 \cf0 Sub StockInfoFormulator()\
Dim ws As Worksheet\
Dim tickerSymbol As String\
Dim yearlyChange As Long\
Dim percentChange As Double\
Dim stockVolume As Long\
 \
Dim summaryTableRow As Long\
\
tickerSymbol = " "\
yearlyChange = 0\
percentChange = 0\
stockVolume = 0\
summaryTableRow = 2\
\
For Each ws In Worksheets\
\
    ws.Cells(1, 9).Value = "Ticker"\
    ws.Cells(1, 10).Value = "Yearly Change"\
    ws.Cells(1, 11).Value = "Percent Change"\
    ws.Cells(1, 12).Value = "Total Stock Volume"\
    \
MaxRow = ws.Cells(Rows.Count, 1).End(xlUp).Row\
MaxColumn = ws.Cells(1, Columns.Count).End(xlToLeft).Column\
\
stockOpen = ws.Cells(2, 3).Value\
                    \
    For i = 2 To MaxRow\
        \
        \
            If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value And ws.Cells(i, 3).Value <> 0 Then\
            tickerSymbol = ws.Cells(i, 1).Value\
            stockClose = ws.Cells(i, 6).Value\
            yearlyChange = stockClose - stockOpen\
            percentChange = (yearlyChange) / stockOpen\
            stockVolume = stockVolume + ws.Cells(i, 7).Value\
            'stockOpen = ws.Cells(i, 3).Value\
            \
\
  \
            ws.Range("I" & summaryTableRow).Value = tickerSymbol\
            ws.Range("L" & summaryTableRow).Value = stockVolume\
            ws.Range("J" & summaryTableRow).Value = yearlyChange\
            ws.Range("K" & summaryTableRow).Value = percentChange\
            \
            summaryTableRow = summaryTableRow + 1\
            \
            stockVolume = 0\
            yearlyChange = 0\
            percentChange = 0\
            stockClose = 0\
         \
        End If\
    Next i\
   \
    For i = 2 To MaxRow\
        \
            ws.Cells(i, 11).Style = "Percent"\
    Next i\
    For i = 2 To MaxRow\
    If (ws.Cells(i, 10).Value > 0) Then\
                    ws.Cells(i, 10).Interior.ColorIndex = 4\
                ElseIf (ws.Cells(i, 10).Value <= 0) Then\
                    ws.Cells(i, 10).Interior.ColorIndex = 3\
                     summaryTableRow = summaryTableRow + 1\
                End If\
    \
    Next i\
\
    \
Next ws\
\
End Sub\
\
\
\
\
\
}