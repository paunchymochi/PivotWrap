VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsPtField"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private p_field_name As String
Private p_source_field_name As String
Private p_orientation As XlPivotFieldOrientation
Private p_position As Long

Private p_rowcol_sort_order As XlSortOrder
Private p_rowcol_sort_by As String

Private p_rowcol_filter_type As XlPivotFilterType
Private p_rowcol_filter_by_field_name As String
Private p_rowcol_filter_value1 As String
Private p_rowcol_filter_value2 As String

Private p_rowcol_manual_filter_values As Variant

Private p_data_function As XlConsolidationFunction
Private p_data_calculation As XlPivotFieldCalculation
Private p_data_number_format As String

Private p_repeat_labels As XlPivotFieldRepeatLabels

Private p_current_page As Variant


Sub init_row(field_name As String, _
        position As Long, _
        Optional sort_by As String, _
        Optional sort_order As XlSortOrder = xlAscending, _
        Optional filter_type As XlPivotFilterType, _
        Optional filter_by_field_name As String, _
        Optional filter_value1 As String, _
        Optional filter_value2 As String, _
        Optional manual_filter_values As Variant = "")
    ' Initialize Row Field
        
    init_rowcol field_name, xlRowField, position, _
            sort_by, sort_order, filter_type, filter_by_field_name, _
            filter_value1, filter_value2, manual_filter_values
    
End Sub


Sub init_col(field_name As String, _
        position As Long, _
        Optional sort_by As String, _
        Optional sort_order As XlSortOrder = xlAscending, _
        Optional filter_type As XlPivotFilterType, _
        Optional filter_by_field_name As String, _
        Optional filter_value1 As String, _
        Optional filter_value2 As String, _
        Optional manual_filter_values As Variant = "")
    ' Initialize Column Field
    
    init_rowcol field_name, xlColumnField, position, _
            sort_by, sort_order, filter_type, filter_by_field_name, _
            filter_value1, filter_value2, manual_filter_values
    
End Sub


Sub init_data(field_name As String, _
        source_field_name As String, _
        position As Long, _
        Optional data_function As XlConsolidationFunction = xlSum, _
        Optional data_calculation As XlPivotFieldCalculation = xlNoAdditionalCalculation, _
        Optional data_number_format As String = "General")
    ' Initialize Data Field
    
    init field_name, source_field_name, xlDataField, position
    
    set_data_function data_function
    set_data_calculation data_calculation
    set_data_number_format data_number_format

End Sub


Sub init_page(field_name As String, _
        position As Long, _
        Optional current_page As Variant = "(All)")
    ' Initialize Page Field
    
    init field_name, field_name, xlPageField, position
    
    If IsMissing(current_page) Or current_page = "" Then
        set_current_page "(All)"
    Else
        set_current_page current_page
    End If
    
End Sub


Private Sub init_rowcol(field_name As String, _
        orientation As XlPivotFieldOrientation, _
        position As Long, _
        Optional sort_by As String, _
        Optional sort_order As XlSortOrder = xlAscending, _
        Optional filter_type As XlPivotFilterType, _
        Optional filter_by_field_name As String, _
        Optional filter_value1 As String, _
        Optional filter_value2 As String, _
        Optional manual_filter_values As Variant = "")
    ' Shared code for initializing Row and Column Fields
        
    init field_name, field_name, orientation, position
    
    set_rowcol_sort_by_field_name sort_by
    set_rowcol_sort_order sort_order
    
    set_rowcol_filter_type filter_type
    set_rowcol_filter_by_field_name filter_by_field_name
    set_rowcol_filter_value1 filter_value1
    set_rowcol_filter_value2 filter_value2
    
    If IsMissing(manual_filter_values) Then
        set_rowcol_manual_filter_values ""
    Else
        set_rowcol_manual_filter_values manual_filter_values
    End If
    
End Sub


Private Sub init(field_name As String, _
        source_field_name As String, _
        orientation As XlPivotFieldOrientation, _
        position As Long)
    ' Shared code for initializing all 4 field orientations
    
    set_field_name field_name
    set_source_field_name source_field_name
    set_orientation orientation
    set_position position
    
End Sub


Public Function get_orientation() As XlPivotFieldOrientation
    get_orientation = p_orientation
End Function


Private Function set_orientation(orientation As XlPivotFieldOrientation) As Boolean
    Dim allowed_values As Variant
    Dim i As Long
    
    allowed_values = Array(XlPivotFieldOrientation.xlRowField, _
            XlPivotFieldOrientation.xlColumnField, _
            XlPivotFieldOrientation.xlPageField, _
            XlPivotFieldOrientation.xlDataField, _
            XlPivotFieldOrientation.xlHidden)
    
    For i = LBound(allowed_values) To UBound(allowed_values)
        If orientation = allowed_values(i) Then
            p_orientation = orientation
            set_orientation = True
            Exit Function
        End If
    Next i
    
    MsgBox "set_orientation failed. Argument orientation must be a member of XlPivotFieldOrientation."
    set_orientation = False
    
End Function


Public Function get_field_name() As String
    get_field_name = p_field_name
End Function


Private Function set_field_name(field_name As String) As Boolean
    If field_name = vbNullString Then
        MsgBox "Field Name cannot be empty"
        set_field_name = False
        Exit Function
    End If
    
    p_field_name = field_name
    set_field_name = True
    
End Function


Public Function get_source_field_name() As String
    get_source_field_name = p_source_field_name
End Function


Private Function set_source_field_name(source_field_name As String) As Boolean
    If source_field_name = vbNullString Then
        MsgBox "Source Field Name cannot be empty"
        set_source_field_name = False
        Exit Function
    End If
    
    p_source_field_name = source_field_name
    set_source_field_name = True
    
End Function


Public Function get_position() As Long
    get_position = p_position
End Function


Public Function set_position(position As Long) As Long
    If position < 1 Then
        MsgBox "set_position failed. Argument position must be greater than or equal to 1."
        set_position = False
    Else
        p_position = position
        set_position = True
    End If
End Function


Public Function get_rowcol_sort_order() As XlSortOrder
    get_rowcol_sort_order = p_rowcol_sort_order
End Function


Public Function set_rowcol_sort_order(sort_order As XlSortOrder) As Boolean
    If sort_order = xlAscending Or sort_order = xlDescending Then
        p_rowcol_sort_order = sort_order
    Else
        MsgBox "set_rowcol_sort_order failed. Argument sort_order must be a member of XlSortOrder."
        set_rowcol_sort_order = False
        Exit Function
    End If
    
    set_rowcol_sort_order = True
    
End Function


Public Function get_rowcol_sort_by_field_name() As String
    get_rowcol_sort_by_field_name = p_rowcol_sort_by
End Function


Public Function set_rowcol_sort_by_field_name(Optional field_name As String) As Boolean

    ' if sort_by is vbnullstring then set sort_by to source_field_name
    If field_name = vbNullString Then
        p_rowcol_sort_by = p_source_field_name
    Else
        p_rowcol_sort_by = field_name
    End If
    
    set_rowcol_sort_by_field_name = True
    
End Function


Public Function set_rowcol_sort(sort_order As XlSortOrder, Optional sort_by_field_name As String) As Boolean
    set_rowcol_sort = set_rowcol_sort_by_field_name(sort_by_field_name) And set_rowcol_sort_order(sort_order)
End Function


Public Function get_rowcol_filter_type() As XlPivotFilterType
    get_rowcol_filter_type = p_rowcol_filter_type
End Function


Public Function set_rowcol_filter_type(filter_type As XlPivotFilterType) As Boolean
    p_rowcol_filter_type = filter_type
    set_rowcol_filter_type = True
End Function


Public Function get_rowcol_filter_by_field_name() As String
    get_rowcol_filter_by_field_name = p_rowcol_filter_by_field_name
End Function


Public Function set_rowcol_filter_by_field_name(field_name As String) As Boolean
    p_rowcol_filter_by_field_name = field_name
    set_rowcol_filter_by_field_name = True
End Function


Public Function get_rowcol_filter_value1() As String
    get_rowcol_filter_value1 = p_rowcol_filter_value1
End Function


Public Function set_rowcol_filter_value1(val As String) As Boolean
    p_rowcol_filter_value1 = val
    set_rowcol_filter_value1 = True
End Function


Public Function get_rowcol_filter_value2() As String
    get_rowcol_filter_value2 = p_rowcol_filter_value2
End Function


Public Function set_rowcol_filter_value2(val As String) As Boolean
    p_rowcol_filter_value2 = val
    set_rowcol_filter_value2 = True
End Function


Public Function set_rowcol_filter(filter_type As XlPivotFilterType, _
        filter_value1 As String, _
        Optional filter_by_field_name As String, _
        Optional filter_value2 As String) As Boolean
            
    Dim ans As Boolean
    
    ans = True
    
    ans = ans And set_rowcol_filter_type(filter_type)
    ans = ans And set_rowcol_filter_by_field_name(filter_by_field_name)
    ans = ans And set_rowcol_filter_value1(filter_value1)
    ans = ans And set_rowcol_filter_value2(filter_value2)
    
    set_rowcol_filter = ans
    
End Function


Public Function get_rowcol_manual_filter_values() As Variant
    get_rowcol_manual_filter_values = p_rowcol_manual_filter_values
End Function


Public Function set_rowcol_manual_filter_values(filter_values As Variant) As Boolean
    Dim arr() As String
    Dim i As Long
    
    If IsArray(filter_values) Then
        ReDim arr(LBound(filter_values) To UBound(filter_values)) As String
        On Error GoTo Err:
            For i = LBound(arr) To UBound(arr)
                arr(i) = CStr(filter_values(i))
            Next i
        On Error GoTo 0
        p_rowcol_manual_filter_values = arr
    Else
        On Error GoTo Err:
            p_rowcol_manual_filter_values = CStr(filter_values)
        On Error GoTo 0
    End If
    
    set_rowcol_manual_filter_values = True
    Exit Function
    
Err:
    On Error GoTo 0
    MsgBox "set_rowcol_manual_filter_values failed. Argument filter_values must be a string or an array of strings"
    set_rowcol_manual_filter_values = False
    
End Function


Public Function get_data_function() As XlConsolidationFunction
    get_data_function = p_data_function
End Function


Public Function set_data_function(data_function As XlConsolidationFunction) As Boolean
    p_data_function = data_function
    set_data_function = True
End Function


Public Function get_data_calculation() As XlPivotFieldCalculation
    get_data_calculation = p_data_calculation
End Function


Public Function set_data_calculation(data_calculation As XlPivotFieldCalculation) As Boolean
    p_data_calculation = data_calculation
    set_data_calculation = True
End Function


Public Function get_data_number_format() As String
    get_data_number_format = p_data_number_format
End Function


Public Function set_data_number_format(number_format As String) As Boolean
    p_data_number_format = number_format
    set_data_number_format = True
End Function


Public Function get_repeat_labels() As XlPivotFieldRepeatLabels
    get_repeat_labels = p_repeat_labels
End Function


Public Function set_repeat_labels(repeat_labels As XlPivotFieldRepeatLabels) As Boolean
    p_repeat_labels = repeat_labels
    set_repeat_labels = True
End Function


Public Function get_current_page() As Variant
    get_current_page = p_current_page
End Function


Public Function set_current_page(current_page As Variant) As Boolean
    Dim pages() As String
    Dim i As Long
    
    If IsArray(current_page) Then
        ReDim pages(LBound(current_page) To UBound(current_page)) As String
        For i = LBound(pages) To UBound(pages)
            pages(i) = CStr(current_page(i))
        Next i
        p_current_page = pages
    Else
        On Error GoTo Err:
        p_current_page = CStr(current_page)
        On Error GoTo 0
    End If
        
    set_current_page = True
    Exit Function
    
Err:
    On Error GoTo 0
    MsgBox "set_current_page failed. Argument current_page must be a string or array of strings."
    set_current_page = False
    
End Function

