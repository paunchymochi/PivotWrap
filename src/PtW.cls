VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ptw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private p_pivot_cache As PivotCache
Private p_source_workbook As Workbook
Private p_source_data As Range

Private p_pivot_tables_dict As Scripting.Dictionary

Private p_row_fields_dict As Scripting.Dictionary
Private p_column_fields_dict As Scripting.Dictionary
Private p_data_fields_dict As Scripting.Dictionary
Private p_page_fields_dict As Scripting.Dictionary

Private p_datapivot_field_orientation As XlPivotFieldOrientation
Private p_datapivot_field_position As Long

Private p_pt_row_grand_visible As Boolean
Private p_pt_column_grand_visible As Boolean

Private p_pt_error_string As String
Private p_pt_null_string As String
Private p_pt_display_error_string As Boolean
Private p_pt_display_null_string As Boolean

Private p_title As String

' === Public Property Accessors for Dictionaries ===
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


Public Property Get pivot_tables() As Scripting.Dictionary
    Set pivot_tables = p_pivot_tables_dict
End Property
' =====


' === Public Get Property Functions ===
Public Function get_source_data() As Range
    Set get_source_data = p_source_data
End Function


Public Function get_source_workbook() As Workbook
    Set get_source_workbook = p_source_workbook
End Function
 

Public Function get_pivot_cache() As PivotCache
    Set get_pivot_cache = p_pivot_cache
End Function


Public Function get_datapivot_field_orientation() As XlPivotFieldOrientation
    get_datapivot_field_orientation = p_datapivot_field_orientation
End Function


Public Function get_datapivot_field_position() As Long
    get_datapivot_field_position = p_datapivot_field_position
End Function


Public Function get_pt_row_grand_visible() As Boolean
    get_pt_row_grand_visible = p_pt_row_grand_visible
End Function


Public Function get_pt_column_grand_visible() As Boolean
    get_pt_column_grand_visible = p_pt_column_grand_visible
End Function


Public Function get_pt_error_string() As String
    get_pt_error_string = p_pt_error_string
End Function


Public Function get_pt_null_string() As String
    get_pt_null_string = p_pt_null_string
End Function


Public Function get_pt_display_error_string() As Boolean
    get_pt_display_error_string = p_pt_display_error_string
End Function


Public Function get_pt_display_null_string() As Boolean
    get_pt_display_null_string = p_pt_display_null_string
End Function
' =====


' === Immediate Window-Friendly Print Functions ===
Public Function print_column_names() As String
    ' Print all column names (i.e. available pivot table fields) from data source
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
    ' Print all added row fields
    print_row_fields = print_dict(row_fields, "Row")
End Function


Public Function print_column_fields() As String
    ' Print all added column fields
    print_column_fields = print_dict(column_fields, "Column")
End Function


Public Function print_data_fields() As String
    ' Print all added data fields
    print_data_fields = print_dict(data_fields, "Data")
End Function


Public Function print_page_fields() As String
    ' Print all added page fields
    print_page_fields = print_dict(page_fields, "Page")
End Function

Public Function print_all_fields() As String
    print_row_fields
    print_column_fields
    print_data_fields
    print_page_fields
End Function


Private Function print_dict(dict As Scripting.Dictionary, field_orientation As String) As String
    ' Print field information in Immediate window
    Dim key As Variant
    Dim s As String
    Dim s2 As String
    Dim field As PtWField
    
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
' =====


Public Function get_column_names() As String()
    Dim strings() As String
    Dim i As Long
    
    ReDim strings(0 To p_source_data.Columns.Count - 1) As String
    For i = 0 To UBound(strings)
        strings(i) = p_source_data.Cells(1, 1 + i).Value
    Next i
    
    get_column_names = strings
    
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


Public Sub init(ByVal source_data As Range)

    ' Initialize class using source data
    Set source_data = validate_source_data(source_data)
    
    If source_data Is Nothing Then
        MsgBox "Source Data not valid"
        Exit Sub
    End If
    
    Set p_source_data = source_data
    Set p_source_workbook = source_data.Parent.Parent
    Set p_pivot_cache = create_pivot_cache(source_data)
    init_dicts
    setup_defaults
End Sub


Public Sub init_with_pivot_cache(pivot_cache As PivotCache)
    ' Initialize class using existing pivot cache
    Assert False
End Sub


Public Sub init_with_pivot_table(pivot_table As PivotTable)
    ' Initialize class using existing pivot table
    Assert False
End Sub


Private Sub init_dicts()
    ' Create empty dictionaries
    Set p_row_fields_dict = New Scripting.Dictionary
    Set p_column_fields_dict = New Scripting.Dictionary
    Set p_data_fields_dict = New Scripting.Dictionary
    Set p_page_fields_dict = New Scripting.Dictionary
    Set p_pivot_tables_dict = New Scripting.Dictionary
End Sub


Private Sub setup_defaults()
    ' Setup default values, called from init
    
    'Display empty and error strings
    p_pt_display_null_string = True
    p_pt_null_string = vbNullString
    p_pt_display_error_string = False
    p_pt_error_string = vbNullString
    
    'Grand Totals
    p_pt_row_grand_visible = True
    p_pt_column_grand_visible = True
    
    
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


Private Function get_position_array(ByRef field_names As Variant, _
        ByRef positions As Variant, _
        ByRef field_orientation_name As String) As Long()
    
    ' Validates positions variable as a number or array of numbers
    ' field_orientation_name argument is for error message only
    ' Returns validated array of numbers, or empty array if not valid
    ' Used in add_*_fields functions
        
    Dim position_array() As Long
    Dim use_positions As Boolean
    Dim err_msg As String
    Dim i As Long
    
    err_msg = vbNullString
    
    ' Convert parameters to array
    If Not IsArray(field_names) Then field_names = Array(field_names)
    If Not IsArray(positions) Then positions = Array(positions)
    
    ' Make sure positions is not missing and has same number of elements as field_names
    If Not IsMissing(positions(0)) Then
        If UBound(positions) <> UBound(field_names) Then ' Array lengths do not match
            err_msg = err_msg & field_orientation_name & _
                    " : positions array must have the same length as field_names array" & _
                    vbNewLine
            use_positions = False
        Else
            ' Make sure each position element is numeric
            If is_array_numeric(positions) Then
                use_positions = True ' positions array is valid
            Else
                err_msg = err_msg & field_orientation_name & _
                        " : Each element in positions array must be numeric" & _
                        vbNewLine
            End If
        End If
    Else ' missing positions
        use_positions = False
    End If
    
    If Not err_msg = vbNullString Then ' positions validation failed, return empty array
        ReDim position_array(-1 To -1) As Long
        MsgBox err_msg
        get_position_array = position_array
        Exit Function
    End If
    
    ReDim position_array(0 To UBound(field_names)) As Long
    
    ' Use positions array if use_position is True, otherwise use i
    If use_positions Then ' valid positions
        For i = 0 To UBound(position_array)
            position_array(i) = positions(i)
        Next i
    Else ' missing positions
        For i = 0 To UBound(position_array)
            position_array(i) = i + 1
        Next i
    End If
    
    get_position_array = position_array
    
End Function


Private Function is_array_numeric(ByRef arr As Variant) As Boolean
    ' Checks if all elements in arr are numeric
    
    Dim i As Long
    
    For i = LBound(arr) To UBound(arr)
        If Not IsNumeric(arr(i)) Then
            is_array_numeric = False
            Exit Function
        End If
    Next i
    
    is_array_numeric = True
    
End Function


Private Function is_array_unique_strings(ByRef arr As Variant) As Boolean
    ' Checks if all elements in arr are unique
    ' Comparisons are case-insensitive
    
    Dim i As Long
    Dim j As Long
    Dim str_i As String
    Dim str_j As String
    
    For i = LBound(arr) To UBound(arr) - 1
        str_i = LCase(CStr(arr(i)))
        For j = i + 1 To UBound(arr)
            str_j = LCase(CStr(arr(j)))
            If str_i = str_j Then
                is_array_unique_strings = False
                Exit Function
            End If
        Next j
    Next i
    
    is_array_unique_strings = True
    
End Function


Private Function get_key_array(ByRef field_names As Variant, _
        ByRef keys As Variant, _
        ByRef field_orientation_name As String) As String()
    
    ' Validates keys variable as a string or array of strings
    ' each element in keys must be unique
    ' field_orientation_name argument is for error message only
    ' Returns validated array of strings, or empty array if not valid
    ' Used in add_*_fields functions
    
    Dim key_array() As String
    Dim use_keys As Boolean
    Dim err_msg As String
    Dim i As Long
    Dim j As Long
    
    err_msg = vbNullString
    
    ' Convert parameters to array
    If Not IsArray(field_names) Then field_names = Array(field_names)
    If Not IsArray(keys) Then keys = Array(keys)
    
    ' Make sure keys is not missing and has same number of elements as field_names
    If Not IsMissing(keys(0)) Then
        If UBound(keys) <> UBound(field_names) Then ' Array lengths do not match
            err_msg = err_msg & field_orientation_name & _
                    " : keys array must have the same length as field_names array" & _
                    vbNewLine
            use_keys = False
        Else
            ' Make sure each key is unique
            If is_array_unique_strings(keys) Then
                use_keys = True ' keys array is valid
            Else
                err_msg = err_msg & field_orientation_name & _
                        " : Each key element in keys array must be unique" & _
                        vbNewLine
                use_keys = False
            End If
        End If
    Else ' missing keys
        use_keys = False
    End If
    
    If Not err_msg = vbNullString Then ' keys validation failed, return empty array
        ReDim key_array(-1 To -1) As String
        MsgBox err_msg
        get_key_array = key_array
        Exit Function
    End If
    
    ReDim key_array(0 To UBound(field_names)) As String
    
    ' Use keys array is use_key is True, other use field_names
    If use_keys Then ' valid keys
        For i = 0 To UBound(key_array)
            key_array(i) = keys(i)
        Next i
    Else ' missing keys
        For i = 0 To UBound(key_array)
            key_array(i) = field_names(i)
        Next i
    End If
    
    get_key_array = key_array
    
End Function


Public Function add_row_fields(field_names As Variant, _
        Optional keys As Variant, _
        Optional positions As Variant) As PtWField()
    ' Add multiple row fields by providing an array of field_names
    ' returns array of PtWFields
    
    add_row_fields = add_rowcolpage_fields(field_names, keys, positions, xlRowField)
    
End Function


Public Function add_column_fields(field_names As Variant, _
        Optional keys As Variant, _
        Optional positions As Variant) As PtWField()
    ' Add multiple column fields by providing an array of field_names
    ' Returns array of PtWFields
    
    add_column_fields = add_rowcolpage_fields(field_names, keys, positions, xlColumnField)
    
End Function


Public Function add_page_fields(field_names As Variant, _
        Optional keys As Variant, _
        Optional positions As Variant) As PtWField()
    ' Add multiple column fields by providing an array of field_names
    ' Returns array of PtWFields
    
    add_page_fields = add_rowcolpage_fields(field_names, keys, positions, xlPageField)
    
End Function


Public Function add_row_field(field_name As String, _
        position As Long, _
        Optional key As String, _
        Optional sort_by As String, _
        Optional sort_order As XlSortOrder = xlAscending, _
        Optional filter_type As XlPivotFilterType, _
        Optional filter_by_field_name As String, _
        Optional filter_value1 As String, _
        Optional filter_value2 As String, _
        Optional manual_filter_values As Variant, _
        Optional subtotal_visible As Boolean = True, _
        Optional subtotal_location As XlSubtototalLocationType = xlAtTop, _
        Optional blank_line_insert As Boolean = False, _
        Optional show_all_items As Boolean = False, _
        Optional repeat_labels As Boolean = False _
        ) As PtWField
    ' Add new row field and store it in p_row_fields_dict
    ' Return the new row field
    
    If key = vbNullString Then
        key = field_name
    End If
        
    If Not validate_field_name_inputs(field_name, field_name, key, p_row_fields_dict) Then
        Exit Function
    End If
    
    Dim ptfield As PtWField
    Set ptfield = New PtWField
    
    ptfield.init_row _
            field_name, _
            position, _
            sort_by, _
            sort_order, _
            filter_type, _
            filter_by_field_name, _
            filter_value1, _
            filter_value2, _
            manual_filter_values, _
            subtotal_visible, _
            subtotal_location, _
            blank_line_insert, _
            show_all_items, _
            repeat_labels

    p_row_fields_dict.Add key, ptfield
    Set add_row_field = ptfield
    
End Function


Public Function add_column_field(field_name As String, _
        position As Long, _
        Optional key As String, _
        Optional sort_by As String, _
        Optional sort_order As XlSortOrder = xlAscending, _
        Optional filter_type As XlPivotFilterType, _
        Optional filter_by_field_name As String, _
        Optional filter_value1 As String, _
        Optional filter_value2 As String, _
        Optional manual_filter_values As Variant, _
        Optional subtotal_visible As Boolean = True, _
        Optional subtotal_location As XlSubtototalLocationType = xlAtTop, _
        Optional blank_line_insert As Boolean = False, _
        Optional show_all_items As Boolean = False, _
        Optional repeat_labels As Boolean = False _
        ) As PtWField
    ' Add new column field and store it in p_row_fields_dict
    ' Return the new column field
    
    If key = vbNullString Then
        key = field_name
    End If
    
    If Not validate_field_name_inputs(field_name, field_name, key, p_column_fields_dict) Then
        Exit Function
    End If
    
    Dim ptfield As PtWField
    Set ptfield = New PtWField
    
    ptfield.init_col _
            field_name, _
            position, _
            sort_by, _
            sort_order, _
            filter_type, _
            filter_by_field_name, _
            filter_value1, _
            filter_value2, _
            manual_filter_values, _
            subtotal_visible, _
            subtotal_location, _
            blank_line_insert, _
            show_all_items, _
            repeat_labels
    
    p_column_fields_dict.Add key, ptfield
    Set add_column_field = ptfield
    
End Function


Public Function add_data_field(field_name As String, _
        source_field_name As String, _
        position As Long, _
        Optional key As String, _
        Optional data_function As XlConsolidationFunction, _
        Optional data_calculation As XlPivotFieldCalculation, _
        Optional data_number_format As String) As PtWField
    ' Add new data field and store it in p_data_fields_dict
    ' Return the new data field

    If key = vbNullString Then
        key = field_name
    End If
    
    If Not validate_field_name_inputs(field_name, source_field_name, key, p_data_fields_dict) Then
        Exit Function
    End If
    
    Dim ptfield As PtWField
    Set ptfield = New PtWField
    
    ptfield.init_data field_name, _
            source_field_name, _
            position, _
            data_function, _
            data_calculation, _
            data_number_format
    
    p_data_fields_dict.Add key, ptfield
    Set add_data_field = ptfield

End Function


Public Function add_page_field(field_name As String, _
        position As Long, _
        Optional key As String, _
        Optional current_page As String) As PtWField
    ' Add new page field and store it in p_page_fields_dict
    ' Return the new page field
    
    If key = vbNullString Then
        key = field_name
    End If
    
    If Not validate_field_name_inputs(field_name, field_name, key, p_page_fields_dict) Then
        Exit Function
    End If
    
    Dim ptfield As PtWField
    Set ptfield = New PtWField
    
    ptfield.init_page field_name, _
            position, _
            current_page
    
    p_page_fields_dict.Add key, ptfield
    Set add_page_field = ptfield
        
End Function


Public Function set_title(title As String) As Boolean
    ' Set title of pivot table
    
    p_title = title
    set_title = True
    
End Function


Public Function set_datapivot_field(data_orientation As XlPivotFieldOrientation, _
        Optional data_position As Long) As Boolean
    ' Set data pivot field orientation & position (i.e. the "Sigma Values" field in PivotTable Fields)
        
    p_datapivot_field_orientation = data_orientation
    p_datapivot_field_position = data_position

    set_datapivot_field = True
    
End Function


Public Function set_grand_totals_visibility(row_visible As Boolean, column_visible As Boolean) As Boolean
    ' Set visibility of row and column grand totals
    p_pt_row_grand_visible = row_visible
    p_pt_column_grand_visible = column_visible
End Function


Public Function set_error_display(display_error As Boolean, Optional error_string As String) As Boolean
    ' Set error display string (i.e. "For error values show" option in PivotTable Options)
    p_pt_display_error_string = display_error
    p_pt_error_string = error_string
    set_error_display = True
End Function


Public Function set_null_display(display_empty As Boolean, Optional null_string As String) As Boolean
    ' Set null display string (i.e. "For empty values show" option in PivotTable Options)
    p_pt_display_null_string = display_empty
    p_pt_null_string = null_string
    set_null_display = True
End Function


Public Function remove_row_field(field_name As String, Optional is_key As Boolean) As PtWField
    ' Removes row field from dictionary
    ' Returns removed row field
    
    Dim pt_field As PtWField
    
    Set pt_field = pop_field(p_row_fields_dict, field_name, is_key)
    Set remove_row_field = pt_field
    
End Function


Public Function remove_column_field(field_name As String, Optional is_key As Boolean) As PtWField
    ' Removes column field from dictionary
    ' Returns removed column field
    
    Dim pt_field As PtWField
    
    Set pt_field = pop_field(p_column_fields_dict, field_name, is_key)
    Set remove_column_field = pt_field
    
End Function


Public Function remove_page_field(field_name As String, Optional is_key As Boolean) As PtWField
    ' Removes page field from dictionary
    ' Returns removed page field
    
    Dim pt_field As PtWField
    
    Set pt_field = pop_field(p_page_fields_dict, field_name, is_key)
    Set remove_page_field = pt_field
    
End Function


Public Function remove_data_field(field_name As String, Optional is_key As Boolean) As PtWField
    ' Removes data field from dictionary
    ' Returns removed data field
    
    Dim pt_field As PtWField
    
    Set pt_field = pop_field(p_data_fields_dict, field_name, is_key)
    Set remove_data_field = pt_field
    
End Function


Private Function add_rowcolpage_fields(ByRef field_names As Variant, _
        ByRef keys As Variant, _
        ByRef positions As Variant, _
        orientation As XlPivotFieldOrientation) As PtWField()
    ' Calls add_row/column_field for each element in field_names
    ' Returns an array of PtWFields
    ' Used in add_row_fields and add_column_fields
    
    Dim i As Long
    Dim position_arr() As Long
    Dim key_arr() As String
    Dim field_orientation As String
    Dim pt_fields() As PtWField
    
    If orientation = xlRowField Then
        field_orientation = "add_row_fields"
    ElseIf orientation = xlColumnField Then
        field_orientation = "add_column_fields"
    ElseIf orientation = xlPageField Then
        field_orientation = "add_page_fields"
    End If
    
    ' Convert arguments to arrays
    If Not IsArray(field_names) Then field_names = Array(field_names)
    If Not IsArray(keys) Then keys = Array(keys)
    If Not IsArray(positions) Then positions = Array(positions)
    
    position_arr = get_position_array(field_names, positions, field_orientation)
    key_arr = get_key_array(field_names, keys, field_orientation)
    
    If UBound(position_arr) = -1 Or UBound(key_arr) = -1 Then ' Invalid arguments provided; abort
        End
    End If
    
    ReDim pt_fields(0 To UBound(field_names)) As PtWField
    
    If orientation = xlRowField Then
        For i = LBound(field_names) To UBound(field_names)
            Set pt_fields(i) = add_row_field(CStr(field_names(i)), position_arr(i), key_arr(i))
        Next i
    ElseIf orientation = xlColumnField Then
        For i = LBound(field_names) To UBound(field_names)
            Set pt_fields(i) = add_column_field(CStr(field_names(i)), position_arr(i), key_arr(i))
        Next i
    ElseIf orientation = xlPageField Then
        For i = LBound(field_names) To UBound(field_names)
            Set pt_fields(i) = add_page_field(CStr(field_names(i)), position_arr(i), key_arr(i))
        Next i
    End If
    
    add_rowcolpage_fields = pt_fields
    
End Function


Private Function pop_field(d As Scripting.Dictionary, field_name As String, is_key As Boolean) As PtWField
    ' Removes dictionary element if field_name is found in dictionary
    ' Searches in dictionary keys if is_key is True, otherwise searchs in get_field_name property of dictionary item
    
    Dim pt_field As PtWField
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


Public Function create(Optional destination_row As Long, _
        Optional destination_column As Long, _
        Optional table_name As String, _
        Optional title As String, _
        Optional destination_worksheet_name As String, _
        Optional destination_workbook As Workbook, _
        Optional row_keys As Variant, _
        Optional column_keys As Variant, _
        Optional data_keys As Variant, _
        Optional page_keys As Variant) As PivotTable
    ' Inserts new pivot table into provided location
    ' Can override fields selection and positions by providing row, column, data, or page keys in parameter
    ' Aborts if location overlaps with an existing pivot table
    ' Returns created pivot table
    
    Dim pt As PivotTable
    Dim destination_range As Range
    Dim destination_worksheet As Worksheet
    
    Dim PtWField_rows() As PtWField
    Dim PtWField_columns() As PtWField
    Dim PtWField_datas() As PtWField
    Dim PtWField_pages() As PtWField
    
    Dim PtField_rows() As PivotField
    Dim PtField_columns() As PivotField
    Dim PtField_datas() As PivotField
    Dim PtField_pages() As PivotField
    
    ' Set Title
    set_title title
    
    ' Set workbook and worksheet
    Set destination_workbook = get_workbook(destination_workbook)
    Set destination_worksheet = get_worksheet(destination_workbook, destination_worksheet_name)
    
    ' Set destination row and column; find next free row/column if row/column is not provided
    validate_destination_range destination_row, destination_column, destination_worksheet
    
    Set destination_range = get_destination_range(Cells(destination_row, destination_column), _
            destination_worksheet_name, destination_workbook)
    
    ' Create empty pivot table
    Set pt = create_pivot_table(p_pivot_cache, _
        destination_range, _
        destination_worksheet_name, _
        destination_workbook, _
        table_name)
    
    If pt Is Nothing Then
        Exit Function
    End If
    
    ' Retrieve PtWFields arrays, sorted by position
    PtWField_rows = populate_input_fields(p_row_fields_dict, row_keys, "Row Fields")
    PtWField_columns = populate_input_fields(p_column_fields_dict, column_keys, "Column Fields")
    PtWField_datas = populate_input_fields(p_data_fields_dict, data_keys, "Data Fields")
    PtWField_pages = populate_input_fields(p_page_fields_dict, page_keys, "Page Fields")
    
    ' Add PtWFields to pivot table, return arrays of PivotField
    PtField_rows = apply_rowcolpage_pt_fields(pt, PtWField_rows)
    PtField_columns = apply_rowcolpage_pt_fields(pt, PtWField_columns)
    PtField_pages = apply_rowcolpage_pt_fields(pt, PtWField_pages)
    PtField_datas = apply_data_pt_fields(pt, PtWField_datas)
    
    ' Set row/column sorting
    apply_rowcol_sort PtWField_rows, PtField_rows
    apply_rowcol_sort PtWField_columns, PtField_columns
    
    ' Set row/column manual filter (i.e. the list box with check marks in the row/column drop-down menu)
    apply_rowcol_manual_filter PtWField_rows, PtField_rows
    apply_rowcol_manual_filter PtWField_columns, PtField_columns
    
    ' Set row/column filters (i.e. top 10, label filter, value filter) in pivot table
    apply_rowcol_filter PtWField_rows, PtField_rows
    apply_rowcol_filter PtWField_columns, PtField_columns
    
    ' Set page filters (i.e. current page(s)) in pivot table
    apply_page_filter PtWField_pages, PtField_pages
    
    ' Set pivot table formatting (i.e. PivotTable Options menu)
    apply_pt_options pt, PtField_rows, PtField_columns
    
    ' Set data pivot field (i.e. Sigma Value in Field List)
    apply_datapivot_field pt, PtField_rows, PtField_columns, PtField_datas
    
    ' Set row/column field options
    apply_rowcol_options PtWField_rows, PtField_rows
    apply_rowcol_options PtWField_columns, PtField_columns
    
    ' Apply title
    apply_title destination_range
    
    ' Add pivot table to dictionary
    update_pivot_table_dict pt
    
    Set create = pt

End Function


Public Function update(destination_row As Long, _
        destination_column As Long, _
        table_name As String, _
        Optional destination_worksheet_name As String, _
        Optional destination_workbook As Workbook, _
        Optional row_keys As Variant, _
        Optional column_keys As Variant, _
        Optional data_keys As Variant, _
        Optional page_keys As Variant) As PivotTable
    ' Deletes pivot table with table_name and creates it again
    
    Dim pt As PivotTable
    
    Set destination_workbook = get_workbook(destination_workbook)
    
    Set pt = lookup_pivot_table(destination_workbook, table_name)
    
    If Not (pt Is Nothing) Then ' pt with table_name found; delete pt
        pt.TableRange2.Clear
        Set pt = Nothing
        
        Set p_pivot_cache = create_pivot_cache(p_source_data) ' pivot cache is deleted automatically if no pivot table references it
    End If
    
    ' call create
    create destination_row, destination_column, _
            table_name, destination_worksheet_name, _
            destination_workbook, row_keys, column_keys, _
            data_keys, page_keys

End Function


Private Function update_pivot_table_dict(pt As PivotTable) As Boolean
    Dim table_name As String
    
    table_name = pt.name
    
    If p_pivot_tables_dict.Exists(table_name) Then
        p_pivot_tables_dict.Remove table_name
    End If
    
    p_pivot_tables_dict.Add table_name, pt
    
    update_pivot_table_dict = True
    
End Function


Private Function apply_rowcol_sort(cls_pt_fields() As PtWField, pt_fields() As PivotField) As Boolean
    ' Applies row and column sorting
    ' Called by create()
    
    Dim cls_pt_field As PtWField
    Dim pt_field As PivotField
    Dim i As Long
    
    If UBound(cls_pt_fields) = -1 Then ' empty cls_pt_fields
        apply_rowcol_sort = False
        Exit Function
    End If
    
    For i = LBound(cls_pt_fields) To UBound(cls_pt_fields)
        Set cls_pt_field = cls_pt_fields(i)
        If cls_pt_field.get_rowcol_sort_by_field_name <> vbNullString Then ' sort is specified
            Set pt_field = pt_fields(i)
            pt_field.AutoSort _
                    Order:=cls_pt_field.get_rowcol_sort_order(), _
                    field:=cls_pt_field.get_rowcol_sort_by_field_name()
        End If
    Next i
    
    apply_rowcol_sort = True

End Function


Private Function apply_rowcol_options(cls_pt_fields() As PtWField, pt_fields() As PivotField) As Boolean
    ' Applies row and column options
    ' Called by create()
        
    Dim cls_pt_field As PtWField
    Dim pt_field As PivotField
    Dim i As Long
    
    If UBound(cls_pt_fields) = -1 Then 'empty cls_pt_fields
        apply_rowcol_options = False
        Exit Function
    End If
    
    For i = LBound(cls_pt_fields) To UBound(cls_pt_fields)
        Set cls_pt_field = cls_pt_fields(i)
        Set pt_field = pt_fields(i)
        
        ' Apply subtotal visibility
        pt_field.Subtotals(1) = True 'Set to True first to disable the other 11 subtotals
        pt_field.Subtotals(1) = cls_pt_field.get_rowcol_subtotal_visible()
        
        ' Set subtotal location
        pt_field.LayoutSubtotalLocation = cls_pt_field.get_rowcol_subtotal_location()
        
        ' Apply blank line
        pt_field.LayoutBlankLine = cls_pt_field.get_rowcol_blank_line_insert()
        
        ' Apply repeat labels
        pt_field.RepeatLabels = cls_pt_field.get_rowcol_repeat_labels()
        
        ' Apply show all items
        pt_field.ShowAllItems = cls_pt_field.get_rowcol_show_all_items()
        
    Next i
    
    apply_rowcol_options = True
    
End Function


Private Function apply_rowcol_filter(cls_pt_fields() As PtWField, pt_fields() As PivotField) As Boolean
    ' Applies row and column filters
    ' Called by create()
    
    Dim cls_pt_field As PtWField
    Dim pt_field As PivotField
    Dim i As Long
    
    If UBound(cls_pt_fields) = -1 Then ' empty cls_pt_fields
        apply_rowcol_filter = False
        Exit Function
    End If
    
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
    
    apply_rowcol_filter = True

End Function


Private Function apply_rowcol_manual_filter(cls_pt_fields() As PtWField, pt_fields() As PivotField) As Boolean
    ' Applies manual filter for row/column
    ' Called by create()
    
    Dim cls_pt_field As PtWField
    Dim pt_field As PivotField
    
    Dim pivot_item As PivotItem
    Dim pivot_item_name As String
    
    Dim manual_filter_values As Variant
        
    Dim i As Long
    Dim j As Long
    Dim match_found As Boolean
    
    If UBound(cls_pt_fields) = -1 Then ' empty cls_pt_fields
        apply_rowcol_manual_filter = False
        Exit Function
    End If
    
    For i = LBound(cls_pt_fields) To UBound(cls_pt_fields)
        Set cls_pt_field = cls_pt_fields(i)
        Set pt_field = pt_fields(i)
        
        manual_filter_values = cls_pt_field.get_rowcol_manual_filter_values()
        
        If Not IsArray(manual_filter_values) Then ' convert single string item into array
            manual_filter_values = Array(manual_filter_values)
        End If
        
        If Not manual_filter_values(0) = vbNullString Then ' filter is specified
        
            ' Convert values to lower case
            For j = LBound(manual_filter_values) To UBound(manual_filter_values)
                manual_filter_values(j) = LCase(manual_filter_values(j))
            Next j
            
            Application.Calculation = xlCalculationManual
            Application.ScreenUpdating = False
            pt_field.Parent.ManualUpdate = True
            
            ' Loop through each pivot_item, set visible to True if match_found
            For Each pivot_item In pt_field.PivotItems
                pivot_item_name = LCase(pivot_item.name)
                match_found = False
                
                For j = LBound(manual_filter_values) To UBound(manual_filter_values)
                    If pivot_item_name = manual_filter_values(j) Then
                        match_found = True
                        Exit For
                    End If
                Next j
                
                pivot_item.Visible = match_found
                
            Next pivot_item
            
            Application.Calculation = xlCalculationAutomatic
            Application.ScreenUpdating = True
            pt_field.Parent.ManualUpdate = False
            
        End If
    Next i
    
    apply_rowcol_manual_filter = True
    
    
End Function


Private Function apply_page_filter(cls_pt_fields() As PtWField, pt_fields() As PivotField) As Boolean
    ' Applies page filters
    ' Called by create()
    
    Dim cls_pt_field As PtWField
    Dim pt_field As PivotField
    
    Dim pivot_item As PivotItem
    Dim pivot_item_name As String
    
    Dim current_page As Variant
    Dim i As Long
    Dim j As Long
    Dim match_found As Boolean
    
    If UBound(cls_pt_fields) = -1 Then ' empty cls_pt_fields
        apply_page_filter = False
        Exit Function
    End If
    
    For i = LBound(cls_pt_fields) To UBound(cls_pt_fields)
        Set cls_pt_field = cls_pt_fields(i)
        Set pt_field = pt_fields(i)
        current_page = cls_pt_field.get_current_page()
        
        pt_field.ClearAllFilters
        
        If IsArray(current_page) Then
            ' multiple page items filter
            pt_field.EnableMultiplePageItems = True
            
            Application.Calculation = xlCalculationManual
            Application.ScreenUpdating = False
            pt_field.Parent.ManualUpdate = True
            
            ' Loop through each pivot_item, set visible to True if match_found
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
            
            Application.Calculation = xlCalculationAutomatic
            Application.ScreenUpdating = True
            pt_field.Parent.ManualUpdate = False

        
        ElseIf cls_pt_field.get_current_page() <> vbNullString Then
            ' single page item filter
            On Error Resume Next ' Do nothing if current_page is not found
                pt_field.CurrentPage = current_page
            On Error GoTo 0
        End If
    Next i
    
    apply_page_filter = True
    Exit Function

End Function


Private Function apply_pt_options(pt As PivotTable, row_fields() As PivotField, column_fields() As PivotField) As Boolean
    ' Applies general pivot table formatting (blank spaces, error/null string display, grand totals, subtotals
    ' Called by create()
    
    ' Grand Totals
    pt.RowGrand = p_pt_row_grand_visible
    pt.ColumnGrand = p_pt_column_grand_visible
    
    ' Empty and Null displays
    pt.DisplayErrorString = p_pt_display_error_string
    pt.DisplayNullString = p_pt_display_null_string
    pt.ErrorString = p_pt_error_string
    pt.NullString = p_pt_null_string
    
End Function


Private Function apply_datapivot_field(ByRef pt As PivotTable, _
        row_fields() As PivotField, _
        column_fields() As PivotField, _
        data_fields() As PivotField) As Boolean
    ' Sets data pivot field orientation and position in pt
    ' Called by create()
    
    Dim o As XlPivotFieldOrientation
    Dim p As Long
    Dim last_position As Long
    
    ' Do nothing if data_fields is empty or if data_fields has only 1 element
    If UBound(data_fields) <= 0 Then
        apply_datapivot_field = False
        Exit Function
    End If
    
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
    
    last_position = Application.WorksheetFunction.Max(1, last_position) ' set minimum to 1
    
    If p = 0 Or p > last_position Then ' default to last place
        p = last_position
    End If

    With pt.DataPivotField
        .orientation = o
        .position = p
    End With
    
    apply_datapivot_field = True
    
End Function


Private Function apply_rowcolpage_pt_fields(ByRef pt As PivotTable, ByRef input_fields() As PtWField)
    ' Populates row/column/page fields in pt
    ' Called by create()
    
    Dim i As Long
    Dim output_fields() As PivotField
    
    If UBound(input_fields) = -1 Then ' empty input_fields
        ReDim output_fields(-1 To -1) As PivotField
        apply_rowcolpage_pt_fields = output_fields
        Exit Function
    End If
    
    ReDim output_fields(LBound(input_fields) To UBound(input_fields)) As PivotField
    
    For i = LBound(input_fields) To UBound(input_fields)
        Set output_fields(i) = pt.PivotFields(input_fields(i).get_field_name())
        output_fields(i).orientation = input_fields(i).get_orientation()
        output_fields(i).position = i + 1
    Next i
    
    apply_rowcolpage_pt_fields = output_fields
        
End Function


Private Function apply_data_pt_fields(ByRef pt As PivotTable, ByRef input_fields() As PtWField)
    ' Populates data fields in pt
    ' Called by create()
    
    Dim i As Long
    Dim output_fields() As PivotField
    
    If UBound(input_fields) = -1 Then ' empty input_fields
        ReDim output_fields(-1 To -1) As PivotField
        apply_data_pt_fields = output_fields
        Exit Function
    End If
    
    ReDim output_fields(LBound(input_fields) To UBound(input_fields)) As PivotField
    
    For i = LBound(input_fields) To UBound(input_fields)
        Set output_fields(i) = pt.AddDataField( _
                pt.PivotFields(input_fields(i).get_source_field_name()), _
                input_fields(i).get_field_name(), _
                input_fields(i).get_data_function)
        output_fields(i).NumberFormat = input_fields(i).get_data_number_format()
    Next i
    
    apply_data_pt_fields = output_fields
    
End Function


Private Function apply_title(ByRef destination_range As Range) As Boolean
    ' Puts title at topleft cell of destination range
    ' Called by create()
    
    If p_title = vbNullString Then ' No Title to add
        apply_title = False
        Exit Function
    End If
    
    With destination_range.Resize(1, 1)
        .Value = p_title
        .Font.Bold = True
        .Font.Underline = True
    End With
    
End Function


Private Function populate_input_fields(dict As Scripting.Dictionary, Optional keys As Variant, Optional dict_name As String) As PtWField()
    ' Collects PtWFields from provided dict
    ' Uses all elements in dict if keys is not provided, other select only items with corresponding keys
    ' dict_name is used to MsgBox display purposes only
    ' Returns collected PtWFields array

    Dim arr() As PtWField
    Dim key As Variant
    Dim i As Long
    
    If dict.Count = 0 Then ' No ptFields in dictionary
        ReDim arr(-1 To -1) As PtWField
        populate_input_fields = arr
        Exit Function
    End If
    
    If IsEmpty(keys) Or IsMissing(keys) Then 'use all ptFields in dictionary, and sort by ptFields.position
        ReDim arr(0 To dict.Count - 1) As PtWField
        For i = LBound(arr) To UBound(arr)
            Set arr(i) = dict.Items(i)
        Next i
        sort_input_fields arr
    Else ' select items with provided keys only, ignore ptFields.position
        If Not IsArray(keys) Then
            keys = Array(keys)
        End If
        
        ReDim arr(0 To UBound(keys) - LBound(keys)) As PtWField
        
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


Private Sub sort_input_fields(ByRef input_fields() As PtWField)
    ' Sorts fields based on their position property
    ' The deciding position of a field is not based on its position property, but on its position in the array
    ' Returns sorted array
    
    Dim current As PtWField
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


Private Function field_name_exists(field_name As String) As Boolean
    ' Check if field_name exists in the column headers of the source data
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


Private Function validate_field_name_inputs(field_name As String, _
        source_field_name As String, _
        key As String, _
        dict As Scripting.Dictionary) As Boolean
    ' Checks if field_name is unique in d, if source_field_name is in source data headers, and
    ' if key exists in d
    ' Return True if validations pass, False otherwise
    
    Dim err_msg As String
    
    err_msg = err_msg & val_source_field_name_is_in_data(source_field_name)
    err_msg = err_msg & val_field_name_is_not_duplicate(field_name, dict)
    err_msg = err_msg & val_key_is_not_duplicate(key, dict)
    
    If err_msg = vbNullString Then
        validate_field_name_inputs = True
    Else
        MsgBox err_msg
        validate_field_name_inputs = False
    End If
    
End Function


Private Function val_source_field_name_is_in_data(ByRef source_field_name As String) As String
    Dim names As Variant
    Dim name As Variant
    Dim err_msg As String
    Dim found As Boolean
    Dim source_field_name_lcase As String
    ' Check if source_field_name is in source data headers
    ' Return error message if source_field_name is not found in headers
    ' Return vbnullstring otherwise
    ' Used in validate_field_name_inputs
    
    ' Get headers in source data
    names = get_column_names()
    
    source_field_name_lcase = LCase(source_field_name)
    
    found = False
    For Each name In names
        If source_field_name_lcase = LCase(name) Then
            found = True
            Exit For
        End If
    Next name
    If Not found Then
        err_msg = "source_field_name [" & source_field_name & "] does not exist in source data. " & vbNewLine
    Else
        err_msg = vbNullString
    End If
    
    val_source_field_name_is_in_data = err_msg

End Function


Private Function val_field_name_is_not_duplicate(field_name As String, dict As Scripting.Dictionary) As String
    ' Return error message if field_name is already in dictionary
    ' Return vbnullstring otherwise
    ' Used in validate_field_name_inputs
    
    Dim i As Long
    Dim pt_field As PtWField
    
    For i = 0 To dict.Count - 1
        Set pt_field = dict.Items(i)
        If pt_field.get_field_name() = field_name Then
            val_field_name_is_not_duplicate = "field_name [" & field_name & "] is already used" & vbNewLine
            Exit Function
        End If
    Next i
    
    val_field_name_is_not_duplicate = vbNullString
    
End Function


Private Function val_key_is_not_duplicate(key As String, dict As Scripting.Dictionary) As String
    ' Return error message if key is already in dictionary
    ' Return vbnullstring otherwise
    ' Used in validate_field_name_inputs
    
    If key = vbNullString Then
        val_key_is_not_duplicate = "key cannot be empty string" & vbNewLine
        Exit Function
    End If
    
    If dict.Exists(key) Then
        val_key_is_not_duplicate = "key [" & key & "] is already used" & vbNewLine
        Exit Function
    End If
    
    val_key_is_not_duplicate = vbNullString
    
End Function


Private Function validate_source_data(ByVal source_data As Range) As Range
    ' Checks source_data for empty columns and duplicate headers
    ' Renames duplicated header if found
    ' Resizes range if extra empty columns and rows are selected
    ' Returns trimmed Range, or Nothing if range is not valid
    
    Dim d As Scripting.Dictionary
    Dim header As String
    Dim i As Long
    
    Set d = New Scripting.Dictionary
    
    If source_data.Columns.Count < 2 Then ' Pivot tables must have at least 2 columns
        Exit Function
    End If
    
    If source_data.rows.Count < 2 Then ' Pivot tables must have at least 1 data row
        Exit Function
    End If
    
    ' Use CurrentRegion to eliminate potential empty columns
    Set source_data = source_data.Resize(1, 1).CurrentRegion
    
    ' Check if headers are unique
    For i = 0 To source_data.Columns.Count - 1
        header = source_data.Resize(1, 1).Offset(0, i).Value
        If header = vbNullString Then ' Empty header
            source_data.Value = "Column 1" ' Directly overwrite range with default
        End If
        
        Do While d.Exists(header) ' Increment header name if duplicated
            header = get_incremented_string(header)
            source_data.Resize(1, 1).Offset(0, i).Value = header
        Loop
        
        d.Add header, i
    Next i
    
    Set validate_source_data = source_data
    
End Function


Private Function validate_destination_range( _
        ByRef destination_row As Long, _
        ByRef destination_column As Long, _
        ByRef destination_ws As Worksheet) As Boolean
    ' Destination row and column are optional. If their value is 0 (i.e. not provided),
    ' then find the next free row/column with 1 cell padding
    ' If both row and column are 0, then row is set to 3 while column is the next free column
    
    ' Used in create()
    
    ' Overwrites destination_row and destination_column with validated values
    
    Dim used_rows As Long
    Dim used_columns As Long
    
    With Range(destination_ws.[A1], destination_ws.UsedRange)
        used_rows = .rows.Count
        used_columns = .Columns.Count
    End With
    
    If destination_row = 0 And destination_column = 0 Then
        destination_row = 3
    End If
    
    If destination_row = 0 Then
        destination_row = used_rows + 2
    End If
    
    If destination_column = 0 Then
        destination_column = used_columns + 2
    End If
        
End Function


Private Function create_pivot_cache(source_data As Range) As PivotCache
    ' Creates pivot cache from source_data
    ' Returns pivot cache
    
    Dim pivot_cache As PivotCache
    Dim wb As Workbook
    
    Set wb = source_data.Parent.Parent
    
    Set pivot_cache = wb.PivotCaches.create(SourceType:=xlDatabase, SourceData:=source_data)
    Set create_pivot_cache = pivot_cache
    
End Function


Private Function create_pivot_table(pivot_cache As PivotCache, _
        ByVal destination_range As Range, _
        Optional destination_ws_name As String, _
        Optional destination_wb As Workbook, _
        Optional table_name As String) As PivotTable
    ' Creates empty pivot table
    ' Returns new pivot table
        
    Dim pivot_table As PivotTable
    Dim destination_r1c1 As String
    
    Set destination_wb = get_workbook(destination_wb)

    table_name = get_unique_table_name(destination_wb, table_name)
    
    If is_pivot_table_overlap(destination_wb, destination_range) Then
        MsgBox "Pivot Tables cannot overlap"
        Exit Function
    End If
    
    Set destination_range = get_destination_range(destination_range, destination_ws_name, destination_wb)
    
    ' If title is present, move destination_range down 2 rows
    If Not p_title = vbNullString Then
        Set destination_range = destination_range.Offset(2)
    End If
    
    ' move destination_range down count(page fields) + 1
    Set destination_range = destination_range.Offset(p_page_fields_dict.Count + 1)
    
    destination_r1c1 = range_to_r1c1(destination_range)
    
    table_name = get_unique_table_name(destination_wb, table_name)
    
    Set pivot_table = pivot_cache.createPivotTable(TableDestination:=destination_r1c1, _
            TableName:=table_name)
    
    Set create_pivot_table = pivot_table

End Function


Private Function get_destination_range(destination_range As Range, _
        Optional destination_ws_name As String, _
        Optional destination_wb As Workbook) As Range
    ' Selects destination range from given parameters
            
    Dim ws As Worksheet
    Set destination_wb = get_workbook(destination_wb)
    Set ws = get_worksheet(destination_wb, destination_ws_name)
    
    Set get_destination_range = ws.Range(destination_range.Address)
    
End Function


Private Function get_unique_table_name(wb As Workbook, table_name As String) As String
    ' Finds a unique pivot table name
    ' Table name defaults to "PivotTable" if table_name is not provided
    ' Table name is suffixed with incrementing 2 digit number until a unique one is found
    ' Returns unique pivot table name

    If table_name = vbNullString Then
        table_name = "PivotTable"
    End If
    
    Do While Not (lookup_pivot_table(wb, table_name) Is Nothing)
        table_name = get_incremented_string(table_name)
    Loop
    
    get_unique_table_name = table_name
    
End Function


Private Function get_incremented_string(str As String) As String
    ' Adds/Increments a 2 digit number with a preceding underscore to provided table_name
    
    Dim v As Variant
    
    v = Split(str, "_")
    
    If UBound(v) = 0 Then
        get_incremented_string = str & "_01"
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
        get_incremented_string = s
        Exit Function
    ElseIf suffix = vbNullString Then
        get_incremented_string = str & "01"
        Exit Function
    Else
        get_incremented_string = str & "_01"
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
    ' Converts range to R1C1 address
    ' Pivot Table creation requires an R1C1 reference
    
    Dim s As String
    
    s = "[" & rng.Parent.Parent.name & "]"
    s = s & rng.Parent.name & "!"
    s = s & get_topleft_r1c1(rng.Address)
    
    range_to_r1c1 = s
    
End Function


Private Function lookup_pivot_table(wb As Workbook, table_name As String) As PivotTable
    ' Checks if pivot table name exists in workbook
    
    Dim ws As Worksheet
    Dim pivot_table As PivotTable
    
    For Each ws In wb.Worksheets
        For Each pivot_table In ws.PivotTables
            If pivot_table.name = table_name Then
                Set lookup_pivot_table = pivot_table
                Exit Function
            End If
        Next pivot_table
    Next ws
    
End Function


Private Function is_pivot_table_overlap(wb As Workbook, destination_range As Range) As Boolean
    ' Checks if destination_range is inside of an existing pivot table
    
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
    ' Uses activeworkbook is parameter is not provided
    
    If wb Is Nothing Then
        Set wb = ActiveWorkbook
    End If
    
    Set get_workbook = wb
    
End Function


Private Function get_worksheet(wb As Workbook, ws_name As String) As Worksheet
    ' Uses PivotTable worksheet name if none is provided
    ' Creates worksheet if it doesn't exist
    
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
    ' Checks if worksheet name exists in workbook
    
    Dim ws As Worksheet
    
    For Each ws In wb.Worksheets
        If ws.name = ws_name Then
            worksheet_exists = True
            Exit Function
        End If
    Next ws
    
    worksheet_exists = False
    
End Function

