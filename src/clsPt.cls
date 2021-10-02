VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private p_pivot_cache As PivotCache
Private p_source_workbook As Workbook
Private p_source_data As Range


Sub init(source_workbook As Workbook, source_data As Range)
    Set p_source_workbook = source_workbook
    Set p_source_data = source_data
    Set p_pivot_cache = createPivotCache(source_workbook, source_data)
End Sub

Sub setup()
' rows, columns, data, filters
'
End Sub

Function create()

End Function

Private Function createPivotCache(wb As Workbook, source_data As Range) As PivotCache
    Dim pivot_cache As PivotCache
    
    Set pivot_cache = wb.PivotCaches.create(SourceType:=xlDatabase, sourceData:=source_data)
    Set createPivotCache = pivot_cache
    
End Function

Private Function createPivotTable(pivot_cache As PivotCache, _
        destination_wb As Workbook, destination_ws_name As String, _
        destination_topleft As String, table_name As String) As PivotTable
        
    Dim pivot_table As PivotTable
    Dim table_destination As Range
    Dim ws As Worksheet
    
    Set ws = get_worksheet(destination_wb, destination_ws_name)

    Set table_destination = ws.Range(destination_topleft)
    
    
    Set pivot_table = pivot_cache.createPivotTable(TableDestination:=table_destination, _
            TableName:=table_name)
    
    Set createPivotTable = pivot_table

End Function

Private Function get_worksheet(wb As Workbook, ws_name As String) As Worksheet
    Dim ws As Worksheet
    
    If worksheet_exists(wb, ws_name) Then
        Set get_worksheet = wb.Worksheets(ws_name)
        Exit Function
    End If
    
    Set ws = wb.Worksheets.Add()
    
    ws.Name = ws_name
    
    Set get_worksheet = ws
    
End Function

Private Function worksheet_exists(wb As Workbook, ws_name As String) As Boolean
    Dim ws As Worksheet
    
    For Each ws In wb.Worksheets
        If ws.Name = ws_name Then
            worksheet_exists = True
            Exit Function
        End If
    Next ws
    
    worksheet_exists = False
    
End Function

