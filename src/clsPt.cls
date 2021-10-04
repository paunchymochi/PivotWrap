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

Private p_row_fields_dict As Scripting.Dictionary
Private p_column_fields_dict As Scripting.Dictionary
Private p_data_fields_dict As Scripting.Dictionary
Private p_page_fields_dict As Scripting.Dictionary
Private d As Scripting.Dictionary


Public Property Get row_fields() As Scripting.Dictionary
    Set row_fields = p_row_fields_dict
End Property


Public Property Get column_fields() As Scripting.Dictionary
    Set column_fields = p_column_fields_dict
End Property


Public Property Get data_fields() As Scripting.Dictionary
    Set data_fields = p_data_fields_dict
End Property


Public Property Get page_fields() As Scripting.Dictionary
    Set page_fields = p_page_fields_dict
End Property


Private Function get_column_names() As String()
    Dim strings() As String
    Dim i As Long
    
    ReDim strings(0 To p_source_data.Columns.Count - 1) As String
    For i = 0 To UBound(strings)
        strings(i) = p_source_data.Cells(1, 1 + i).Value
    Next i
    
    get_column_names = strings
    
End Function


Function print_column_names() As String
    Dim names As Variant
    Dim i As Long
    Dim index_str As String
    Dim s As String
    
    names = get_column_names()
    
    s = s & "=== column_names ===" & vbNewLine
    
    For i = LBound(names) To UBound(names)
        If i < 10 Then
            index_str = "0" & CStr(i)
        Else
            index_str = CStr(i)
        End If
        s = s & index_str & " : " & names(i) & vbNewLine
    Next i
    
    s = s & "=== end ==="
    
    Debug.Print (s)
    
    print_column_names = s
    
End Function


Function print_row_fields() As String
    print_row_fields = print_dict(row_fields, "Row")
End Function


Function print_column_fields() As String
    print_column_fields = print_dict(column_fields, "Column")
End Function


Function print_data_fields() As String
    print_data_fields = print_dict(data_fields, "Data")
End Function


Function print_page_fields() As String
    print_page_fields = print_dict(page_fields, "Page")
End Function

Function print_all_fields() As String
    print_row_fields
    print_column_fields
    print_data_fields
    print_page_fields
End Function


Private Function print_dict(dict As Scripting.Dictionary, field_orientation As String) As String
    Dim key As Variant
    Dim s As String
    Dim s2 As String
    Dim field As clsPtField
    
    s = vbNewLine & "=== " & field_orientation & " Fields ===" & vbNewLine
    
    For Each key In dict.Keys
        s2 = field_orientation & " Field: " & key & vbNewLine
        Set field = dict(key)
        s2 = s2 & vbTab & "Position: " & CStr(field.get_position()) & vbNewLine
        s = s & s2
    Next key
    
    s = s & "=== End ==="
    
    Debug.Print (s)
    
    print_dict = s
        
End Function


Sub init(source_workbook As Workbook, source_data As Range)
    Set p_source_workbook = source_workbook
    Set p_source_data = source_data
    Set p_pivot_cache = createPivotCache(source_workbook, source_data)
    init_dicts
End Sub


Sub init_with_pivot_cache(pivot_cache As PivotCache)
    Assert False
End Sub


Sub init_with_pivot_table(pivot_table As PivotTable)
    Assert False
End Sub


Private Sub init_dicts()
    Set p_row_fields_dict = New Scripting.Dictionary
    Set p_column_fields_dict = New Scripting.Dictionary
    Set p_data_fields_dict = New Scripting.Dictionary
    Set p_page_fields_dict = New Scripting.Dictionary
End Sub


Sub setup(Optional ByVal row_fields As Variant, _
        Optional ByVal column_fields As Variant, _
        Optional ByVal data_fields As Variant, _
        Optional ByVal page_fields As Variant)
' rows, columns, data, filters
'

If IsArray(row_fields) = False Then
    row_fields = Array(row_fields)
End If
If IsArray(column_fields) = False Then
    column_fields = Array(column_fields)
End If
If IsArray(data_fields) = False Then
    data_fields = Array(data_fields)
End If
If IsArray(page_fields) = False Then
    page_fields = Array(page_fields)
End If


End Sub


Function add_row_field(field_name As String, _
        position As Long, _
        Optional sort_by As String, _
        Optional sort_order As XlSortOrder, _
        Optional filter_type As XlPivotFilterType, _
        Optional filter_by_field_name As String, _
        Optional filter_value1 As String, _
        Optional filter_value2 As String) As clsPtField
    
    Dim ptfield As clsPtField
    Set ptfield = New clsPtField
    
    ptfield.init_row field_name, _
            position, _
            sort_by, _
            sort_order, _
            filter_type, _
            filter_by_field_name, _
            filter_value1, _
            filter_value2
    
    p_row_fields_dict.Add field_name, ptfield
    Set add_row_field = ptfield
    
End Function


Function add_column_field(field_name As String, _
        position As Long, _
        Optional sort_by As String, _
        Optional sort_order As XlSortOrder, _
        Optional filter_type As XlPivotFilterType, _
        Optional filter_by_field_name As String, _
        Optional filter_value1 As String, _
        Optional filter_value2 As String)
    
    Dim ptfield As clsPtField
    Set ptfield = New clsPtField
    
    ptfield.init_col field_name, _
            position, _
            sort_by, _
            sort_order, _
            filter_type, _
            filter_by_field_name, _
            filter_value1, _
            filter_value2
    
    p_column_fields_dict.Add field_name, ptfield
    Set add_column_field = ptfield
    
End Function


Function add_data_field(field_name As String, _
        position As Long, _
        Optional data_function As XlConsolidationFunction, _
        Optional data_calculation As XlPivotFieldCalculation, _
        Optional data_number_format As String)
    
    Dim ptfield As clsPtField
    Set ptfield = New clsPtField
    
    ptfield.init_data field_name, _
            position, _
            data_function, _
            data_calculation, _
            data_number_format
    
    p_data_fields_dict.Add field_name, ptfield
    Set add_data_field = ptfield

End Function


Function add_page_field(field_name As String, _
        position As Long, _
        Optional current_page As String)
    
    Dim ptfield As clsPtField
    Set ptfield = New clsPtField
    
    ptfield.init_page field_name, _
            position, _
            current_page
    
    p_page_fields_dict.Add field_name, ptfield
    Set add_page_field = ptfield
        
End Function


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

