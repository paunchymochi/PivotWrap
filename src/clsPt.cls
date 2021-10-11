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

Private p_datapivot_field_orientation As XlPivotFieldOrientation
Private p_datapivot_field_position As Long
Private p_row_grand_visible As Boolean
Private p_column_grand_visible As Boolean


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


Function get_source_data() As Range
    Set get_source_data = p_source_data
End Function


Function get_source_workbook() As Workbook
    Set get_source_workbook = p_source_workbook
End Function
 

Function get_pivot_cache() As PivotCache
    Set get_pivot_cache = p_pivot_cache
End Function


Function get_column_names() As String()
    Dim strings() As String
    Dim i As Long
    
    ReDim strings(0 To p_source_data.columns.Count - 1) As String
    For i = 0 To UBound(strings)
        strings(i) = p_source_data.Cells(1, 1 + i).Value
    Next i
    
    get_column_names = strings
    
End Function


Public Function print_column_names() As String
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


Public Function print_row_fields() As String
    print_row_fields = print_dict(row_fields, "Row")
End Function


Public Function print_column_fields() As String
    print_column_fields = print_dict(column_fields, "Column")
End Function


Public Function print_data_fields() As String
    print_data_fields = print_dict(data_fields, "Data")
End Function


Public Function print_page_fields() As String
    print_page_fields = print_dict(page_fields, "Page")
End Function

Public Function print_all_fields() As String
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
    
    For Each key In dict.keys
        s2 = field_orientation & " Field: " & key & vbNewLine
        Set field = dict(key)
        s2 = s2 & vbTab & "Position: " & CStr(field.get_position()) & vbNewLine
        s = s & s2
    Next key
    
    s = s & "=== End ==="
    
    Debug.Print (s)
    
    print_dict = s
        
End Function


Function validate_data_field_names() As Boolean
    Debug.Assert False
End Function


Public Function get_topright_cell(Optional column_offset As Long, Optional at_page_fields As Boolean) As Range
    Debug.Assert False
End Function


Public Function get_bottomleft_cell(Optional row_offset As Long) As Range
    Debug.Assert False
End Function
 

Public Function get_topleft_cell(Optional at_page_fields As Boolean) As Range
    Debug.Assert False
End Function


Public Sub init(source_data As Range)
    Set p_source_workbook = source_data.Parent.Parent
    Set p_source_data = source_data
    Set p_pivot_cache = create_pivot_cache(source_data)
    init_dicts
End Sub


Public Sub init_with_pivot_cache(pivot_cache As PivotCache)
    Assert False
End Sub


Public Sub init_with_pivot_table(pivot_table As PivotTable)
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


Public Function add_row_fields(field_names As Variant, Optional positions As Variant) As clsPtField()
    Debug.Assert False
End Function


Public Function add_row_field(field_name As String, _
        Position As Long, _
        Optional key As String, _
        Optional sort_by As String, _
        Optional sort_order As XlSortOrder, _
        Optional filter_type As XlPivotFilterType, _
        Optional filter_by_field_name As String, _
        Optional filter_value1 As String, _
        Optional filter_value2 As String) As clsPtField
    
    If field_name_exists(field_name) = False Then
        Debug.Print ("[" & field_name & "] is not a found in the source data")
        Exit Function
    End If
    
    Dim ptfield As clsPtField
    Set ptfield = New clsPtField
    
    ptfield.init_row field_name, _
            Position, _
            sort_by, _
            sort_order, _
            filter_type, _
            filter_by_field_name, _
            filter_value1, _
            filter_value2
    
    If key = vbNullString Then
        key = field_name
    End If
    
    p_row_fields_dict.Add key, ptfield
    Set add_row_field = ptfield
    
End Function


Public Function add_column_field(field_name As String, _
        Position As Long, _
        Optional key As String, _
        Optional sort_by As String, _
        Optional sort_order As XlSortOrder, _
        Optional filter_type As XlPivotFilterType, _
        Optional filter_by_field_name As String, _
        Optional filter_value1 As String, _
        Optional filter_value2 As String) As clsPtField
        
    If field_name_exists(field_name) = False Then
        Debug.Print ("[" & field_name & "] is not a found in the source data")
        Exit Function
    End If
    
    Dim ptfield As clsPtField
    Set ptfield = New clsPtField
    
    ptfield.init_col field_name, _
            Position, _
            sort_by, _
            sort_order, _
            filter_type, _
            filter_by_field_name, _
            filter_value1, _
            filter_value2
    
    If key = vbNullString Then
        key = field_name
    End If
    
    p_column_fields_dict.Add key, ptfield
    Set add_column_field = ptfield
    
End Function


Public Function add_data_field(field_name As String, _
        field_name_original As String, _
        Position As Long, _
        Optional key As String, _
        Optional data_function As XlConsolidationFunction, _
        Optional data_calculation As XlPivotFieldCalculation, _
        Optional data_number_format As String) As clsPtField
    
    If field_name_exists(field_name_original) = False Then
        Debug.Print ("[" & field_name_original & "] is not a found in the source data")
        Exit Function
    End If
    
    Dim ptfield As clsPtField
    Set ptfield = New clsPtField
    
    ptfield.init_data field_name, _
            field_name_original, _
            Position, _
            data_function, _
            data_calculation, _
            data_number_format
    
    If key = vbNullString Then
        key = field_name
    End If
    
    p_data_fields_dict.Add key, ptfield
    Set add_data_field = ptfield

End Function


Public Function add_page_field(field_name As String, _
        Position As Long, _
        Optional key As String, _
        Optional current_page As String) As clsPtField
    
    If field_name_exists(field_name) = False Then
        Debug.Print ("[" & field_name & "] is not a found in the source data")
        Exit Function
    End If
    
    Dim ptfield As clsPtField
    Set ptfield = New clsPtField
    
    ptfield.init_page field_name, _
            Position, _
            current_page
    
    If key = vbNullString Then
        key = field_name
    End If
    
    p_page_fields_dict.Add key, ptfield
    Set add_page_field = ptfield
        
End Function


Public Function set_datapivot_field(data_orientation As XlPivotFieldOrientation, _
        Optional data_position As Long) As Boolean
        
    p_datapivot_field_orientation = data_orientation
    p_datapivot_field_position = data_position

    set_datapivot_field = True
    
End Function


Public Function set_grand_totals_visibility(row_visible As Boolean, column_visible As Boolean) As Boolean
    p_row_grand_visible = row_visible
    p_column_grand_visible = column_visible
End Function


Private Function field_name_exists(field_name As String) As Boolean
    Dim names As Variant
    Dim name As Variant
    
    names = get_column_names()
    
    For Each name In names
        If field_name = name Then
            field_name_exists = True
            Exit Function
        End If
    Next name
    
    field_name_exists = False

End Function


Public Function remove_row_field(field_name As String, Optional is_key As Boolean) As clsPtField
    Dim pt_field As clsPtField
    
    Set pt_field = pop_field(p_row_fields_dict, field_name, is_key)
    Set remove_row_field = pt_field
    
End Function


Public Function remove_column_field(field_name As String, Optional is_key As Boolean) As clsPtField
    Dim pt_field As clsPtField
    
    Set pt_field = pop_field(p_column_fields_dict, field_name, is_key)
    Set remove_column_field = pt_field
    
End Function


Public Function remove_page_field(field_name As String, Optional is_key As Boolean) As clsPtField
    Dim pt_field As clsPtField
    
    Set pt_field = pop_field(p_page_fields_dict, field_name, is_key)
    Set remove_page_field = pt_field
    
End Function


Public Function remove_data_field(field_name As String, Optional is_key As Boolean) As clsPtField
    Dim pt_field As clsPtField
    
    Set pt_field = pop_field(p_data_fields_dict, field_name, is_key)
    Set remove_data_field = pt_field
    
End Function


Private Function pop_field(d As Scripting.Dictionary, field_name As String, is_key As Boolean) As clsPtField
    Dim pt_field As clsPtField
    Dim i As Long
    
    For i = 0 To d.Count - 1
        If is_key Then
            If field_name = d.keys(i) Then
                Set pt_field = d.Items(i)
                Set pop_field = pt_field
                d.Remove d.keys(i)
                Exit Function
            End If
        Else
            Set pt_field = d.Items(i)
            If pt_field.get_field_name() = field_name Then
                Set pop_field = pt_field
                d.Remove d.keys(i)
                Exit Function
            End If
        End If
    Next i
    
End Function


Public Function clear_row_fields() As Boolean
    p_row_fields_dict.RemoveAll
    clear_row_fields = True
End Function


Public Function clear_column_fields() As Boolean
    p_column_fields_dict.RemoveAll
    clear_column_fields = True
End Function


Public Function clear_page_fields() As Boolean
    p_page_fields_dict.RemoveAll
    clear_page_fields = True
End Function


Public Function clear_data_fields() As Boolean
    p_data_fields_dict.RemoveAll
    clear_data_fields = True
End Function


Public Function create(destination_row As Long, _
        destination_column As Long, _
        Optional table_name As String, _
        Optional destination_worksheet_name As String, _
        Optional destination_workbook As Workbook, _
        Optional row_keys As Variant, _
        Optional column_keys As Variant, _
        Optional data_keys As Variant, _
        Optional page_keys As Variant) As PivotTable
    
    Dim pt As PivotTable
    Dim destination_range As Range
    Dim destination_worksheet As Worksheet
    
    Dim in_rows() As clsPtField
    Dim in_columns() As clsPtField
    Dim in_datas() As clsPtField
    Dim in_pages() As clsPtField
    
    Dim out_rows() As PivotField
    Dim out_columns() As PivotField
    Dim out_datas() As PivotField
    Dim out_pages() As PivotField
    
    Set destination_workbook = get_workbook(destination_workbook)
    Set destination_worksheet = get_worksheet(destination_workbook, destination_worksheet_name)
    
    Set destination_range = get_destination_range(Cells(destination_row, destination_column), _
            destination_worksheet_name, destination_workbook)

    Set pt = create_pivot_table(p_pivot_cache, _
        destination_range, _
        destination_worksheet_name, _
        destination_workbook, _
        table_name)
    
    If pt Is Nothing Then
        Exit Function
    End If
    
    in_rows = populate_input_fields(p_row_fields_dict, row_keys, "Row Fields")
    in_columns = populate_input_fields(p_column_fields_dict, column_keys, "Column Fields")
    in_datas = populate_input_fields(p_data_fields_dict, data_keys, "Data Fields")
    in_pages = populate_input_fields(p_page_fields_dict, page_keys, "Page Fields")
    
    out_rows = apply_rowcolpage_pt_fields(pt, in_rows)
    out_columns = apply_rowcolpage_pt_fields(pt, in_columns)
    out_pages = apply_rowcolpage_pt_fields(pt, in_pages)
    out_datas = apply_data_pt_fields(pt, in_datas)
    
    apply_rowcol_filter in_rows, out_rows
    apply_rowcol_filter in_columns, out_columns
    
    apply_page_filter in_pages, out_pages
    
    apply_pt_formatting pt, out_rows, out_columns
    
    apply_datapivot_field pt, out_rows, out_columns
    
    Debug.Assert False

End Function


Private Function apply_rowcol_filter(cls_pt_fields() As clsPtField, pt_fields() As PivotField) As Boolean
    Dim cls_pt_field As clsPtField
    Dim pt_field As PivotField
    Dim i As Long
    
    For i = LBound(cls_pt_fields) To UBound(cls_pt_fields)
        Set cls_pt_field = cls_pt_fields(i)
        If cls_pt_field.get_rowcol_filter_type <> 0 Then ' pt field with no filter has filter type value of 0
            Set pt_field = pt_fields(i)
                pt_field.PivotFilters.Add2 _
                        Type:=cls_pt_field.get_rowcol_filter_type(), _
                        DataField:=cls_pt_field.get_rowcol_filter_by_field_name(), _
                        Value1:=cls_pt_field.get_rowcol_filter_value1(), _
                        Value2:=cls_pt_field.get_rowcol_filter_value2()
        End If
    Next i
                        
End Function


Private Function apply_page_filter(cls_pt_fields() As clsPtField, pt_fields() As PivotField) As Boolean
    Dim cls_pt_field As clsPtField
    Dim pt_field As PivotField
    
    Dim pivot_item As PivotItem
    Dim pivot_item_name As String
    
    Dim current_page As Variant
    Dim i As Long
    Dim j As Long
    Dim match_found As Boolean
    
    For i = LBound(cls_pt_fields) To UBound(cls_pt_fields)
        Set cls_pt_field = cls_pt_fields(i)
        Set pt_field = pt_fields(i)
        current_page = cls_pt_field.get_current_page()
        
        pt_field.ClearAllFilters
        
        If IsArray(current_page) Then
            ' multiple page items filter
            pt_field.EnableMultiplePageItems = True
            For Each pivot_item In pt_field.PivotItems
                pivot_item_name = pivot_item.name
                match_found = False
                For j = LBound(current_page) To UBound(current_page)
                    If pivot_item_name = current_page(j) Then
                        match_found = True
                        Exit For
                    End If
                Next j
                pivot_item.Visible = match_found
            Next pivot_item
        
        ElseIf cls_pt_field.get_current_page() <> vbNullString Then
            ' single page item filter
            
            pt_field.CurrentPage = cls_pt_field.get_current_page()
        End If
    Next i

End Function


Private Function apply_pt_formatting(pt As PivotTable, row_fields() As PivotField, column_fields() As PivotField) As Boolean
' blank space, error display, grand totals, subtotals
    
    'Grand Totals
    pt.RowGrand = p_row_grand_visible
    pt.ColumnGrand = p_column_grand_visible
    
    Debug.Assert False
End Function


Private Function populate_input_fields(dict As Scripting.Dictionary, Optional keys As Variant, Optional dict_name As String) As clsPtField()

    Dim arr() As clsPtField
    Dim key As Variant
    Dim i As Long
    
    If IsEmpty(keys) Or IsMissing(keys) Then 'use all ptFields in dictionary, and sort by ptFields.position
        ReDim arr(0 To dict.Count - 1) As clsPtField
        For i = LBound(arr) To UBound(arr)
            Set arr(i) = dict.Items(i)
        Next i
        sort_input_fields arr
    Else ' select items with provided keys only, ignore ptFields.position
        If Not IsArray(keys) Then
            keys = Array(keys)
        End If
        
        ReDim arr(0 To UBound(keys) - LBound(keys)) As clsPtField
        
        On Error GoTo KeyErr:
        For i = LBound(arr) To UBound(arr)
            Set arr(i) = dict(keys(LBound(keys) + i))
        Next i
        On Error GoTo 0
    End If
    
    populate_input_fields = arr
    Exit Function
    
KeyErr:
    MsgBox "Key [" & keys(LBound(keys) + i) & "] not found in dictionary" & dict_name
    Resume Next
        
End Function


Private Sub sort_input_fields(ByRef input_fields() As clsPtField)
    Dim current As clsPtField
    Dim i As Long, j As Long
    
    For i = LBound(input_fields) To UBound(input_fields)
        Set current = input_fields(i)
        
        For j = i - 1 To LBound(input_fields) Step -1
            If current.get_position() < input_fields(j).get_position() Then
                Set input_fields(j + 1) = input_fields(j)
            Else
                Exit For
            End If
        Next j
        Set input_fields(j + 1) = current
    Next i

End Sub


Private Function apply_datapivot_field(ByRef pt As PivotTable, _
        row_fields() As PivotField, _
        column_fields() As PivotField) As Boolean
    
    Dim o As XlPivotFieldOrientation
    Dim p As Long
    Dim last_position As Long
    
    o = p_datapivot_field_orientation
    p = p_datapivot_field_position
    
    If o = 0 Then ' default to column field
        o = xlColumnField
    End If
    
    If o = xlColumnField Then
        last_position = UBound(column_fields) + 2
    ElseIf o = xlRowField Then
        last_position = UBound(row_fields) + 2
    Else
        Debug.Assert False ' Unknown orientation
    End If
    
    If p = 0 Or p > last_position Then ' default to last place
        p = last_position
    End If

    With pt.DataPivotField
        .Orientation = o
        .Position = p
    End With
    
End Function


Private Function apply_rowcolpage_pt_fields(ByRef pt As PivotTable, ByRef input_fields() As clsPtField)
    Dim i As Long
    Dim output_fields() As PivotField
    
    ReDim output_fields(LBound(input_fields) To UBound(input_fields)) As PivotField
    
    For i = LBound(input_fields) To UBound(input_fields)
        Set output_fields(i) = pt.PivotFields(input_fields(i).get_field_name())
        output_fields(i).Orientation = input_fields(i).get_orientation()
        output_fields(i).Position = i + 1
    Next i
    
    apply_rowcolpage_pt_fields = output_fields
        
End Function


Private Function apply_data_pt_fields(ByRef pt As PivotTable, ByRef input_fields() As clsPtField)
    Dim i As Long
    Dim output_fields() As PivotField
    
    ReDim output_fields(LBound(input_fields) To UBound(input_fields)) As PivotField
    
    For i = LBound(input_fields) To UBound(input_fields)
        Set output_fields(i) = pt.AddDataField( _
                pt.PivotFields(input_fields(i).get_field_name_original()), _
                input_fields(i).get_field_name(), _
                input_fields(i).get_data_function)
        output_fields(i).NumberFormat = input_fields(i).get_data_number_format()
    Next i
    
    apply_data_pt_fields = output_fields
    
End Function


Private Function create_pivot_cache(source_data As Range) As PivotCache
    Dim pivot_cache As PivotCache
    Dim wb As Workbook
    
    Set wb = source_data.Parent.Parent
    
    Set pivot_cache = wb.PivotCaches.create(SourceType:=xlDatabase, SourceData:=source_data)
    Set create_pivot_cache = pivot_cache
    
End Function


Private Function create_pivot_table(pivot_cache As PivotCache, _
        destination_range As Range, _
        Optional destination_ws_name As String, _
        Optional destination_wb As Workbook, _
        Optional table_name As String) As PivotTable
        
    Dim pivot_table As PivotTable
    Dim destination_r1c1 As String
    
    Set destination_wb = get_workbook(destination_wb)

    table_name = get_unique_table_name(destination_wb, table_name)
    
    If is_pivot_table_overlap(destination_wb, destination_range) Then
        MsgBox "Pivot Tables cannot overlap"
        Exit Function
    End If
    
    Set destination_range = get_destination_range(destination_range, destination_ws_name, destination_wb)
    
    destination_r1c1 = range_to_r1c1(destination_range)
    
    table_name = get_unique_table_name(destination_wb, table_name)
    
    Set pivot_table = pivot_cache.createPivotTable(TableDestination:=destination_r1c1, _
            TableName:=table_name)
    
    Set create_pivot_table = pivot_table

End Function


Private Function get_destination_range(destination_range As Range, _
        Optional destination_ws_name As String, _
        Optional destination_wb As Workbook) As Range
            
    Dim ws As Worksheet
    Set destination_wb = get_workbook(destination_wb)
    Set ws = get_worksheet(destination_wb, destination_ws_name)
    
    Set get_destination_range = ws.Range(destination_range.Address)
    
End Function


Private Function get_unique_table_name(wb As Workbook, table_name As String) As String

    If table_name = vbNullString Then
        table_name = "PivotTable"
    End If
    
    Do While is_pivot_table_name_exists(wb, table_name) = True
        table_name = get_table_name_suffixed(table_name)
    Loop
    
    get_unique_table_name = table_name
    
End Function


Private Function get_table_name_suffixed(table_name As String) As String
    Dim v As Variant
    
    v = Split(table_name, "_")
    
    If UBound(v) = 0 Then
        get_table_name_suffixed = table_name & "_01"
        Exit Function
    End If
    
    Dim suffix As String
    Dim suffix_num As Long
    
    suffix = v(UBound(v))
    
    If IsNumeric(suffix) Then
        suffix_num = CLng(suffix)
        suffix_num = suffix_num + 1
        suffix = CStr(suffix_num)
        Do While Len(suffix) < 2
            suffix = "0" & suffix
        Loop
        
        Dim i As Long
        Dim s As String
        For i = LBound(v) To UBound(v) - 1
            s = s & v(i)
        Next i
        s = s & "_" & suffix
        get_table_name_suffixed = s
        Exit Function
    ElseIf suffix = vbNullString Then
        get_table_name_suffixed = table_name & "01"
        Exit Function
    Else
        get_table_name_suffixed = table_name & "_01"
        Exit Function
    End If
    
End Function


Private Function get_topleft_r1c1(topleft As String) As String
    ' Converts address string into R1C1 format
    Dim rng As Range
    Dim r As String, c As String
    
    If UCase(topleft) Like "R*C*" Then
        get_topleft_r1c1 = topleft
        Exit Function
    End If
    
    Set rng = Range(topleft)
    r = CStr(rng.Row)
    c = CStr(rng.Column)
    
    get_topleft_r1c1 = "R" & r & "C" & c
    
End Function


Private Function range_to_r1c1(rng As Range) As String
    Dim s As String
    
    s = "[" & rng.Parent.Parent.name & "]"
    s = s & rng.Parent.name & "!"
    s = s & get_topleft_r1c1(rng.Address)
    
    range_to_r1c1 = s
    
End Function


Private Function is_pivot_table_name_exists(wb As Workbook, table_name As String) As Boolean
    Dim ws As Worksheet
    Dim pivot_table As PivotTable
    
    For Each ws In wb.Worksheets
        For Each pivot_table In ws.PivotTables
            If pivot_table.name = table_name Then
                is_pivot_table_name_exists = True
                Exit Function
            End If
        Next pivot_table
    Next ws
    
    is_pivot_table_name_exists = False
    
End Function


Private Function is_pivot_table_overlap(wb As Workbook, destination_range As Range) As Boolean
    Dim pivot_table As PivotTable
    Dim ws As Worksheet
    Dim rng As Range
    
    For Each ws In wb.Worksheets
        If ws.name = destination_range.Parent.name Then
            For Each pivot_table In ws.PivotTables
                On Error GoTo Err:
                    Set rng = pivot_table.TableRange2
                    Debug.Print (ws.name & " : " & pivot_table.TableRange2.Address)
                On Error GoTo 0
                If Not rng Is Nothing Then
                    If Application.Union(rng, destination_range).Address = rng.Address Then
                        is_pivot_table_overlap = True
                        Exit Function
                    End If
                End If
            Next pivot_table
        End If
    Next ws
    
    is_pivot_table_overlap = False
    
    Exit Function
    
Err:
    Set rng = Nothing
    Resume Next
        
End Function


Private Function get_workbook(wb As Variant) As Workbook
    If wb Is Nothing Then
        Set wb = ActiveWorkbook
    End If
    
    Set get_workbook = wb
    
End Function


Private Function get_worksheet(wb As Workbook, ws_name As String) As Worksheet
    Dim ws As Worksheet
    
    If ws_name = vbNullString Then
        ws_name = "PivotTable"
    End If
    
    If worksheet_exists(wb, ws_name) Then
        Set get_worksheet = wb.Worksheets(ws_name)
        Exit Function
    End If
    
    ActiveSheet.Select
    Set ws = wb.Worksheets.Add()
    
    ws.name = ws_name
    
    Set get_worksheet = ws
    
End Function


Private Function worksheet_exists(wb As Workbook, ws_name As String) As Boolean
    Dim ws As Worksheet
    
    For Each ws In wb.Worksheets
        If ws.name = ws_name Then
            worksheet_exists = True
            Exit Function
        End If
    Next ws
    
    worksheet_exists = False
    
End Function

