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

Private p_data_function As XlConsolidationFunction
Private p_data_calculation As XlPivotFieldCalculation
Private p_data_number_format As String

Private p_repeat_labels As XlPivotFieldRepeatLabels

Private p_current_page As String


Sub init_row(field_name As String, _
        position As Long, _
        Optional sort_by As String, _
        Optional sort_order As XlSortOrder, _
        Optional filter_type As XlPivotFilterType, _
        Optional filter_by_field_name As String, _
        Optional filter_value1 As String, _
        Optional filter_value2 As String)
    ' Initialize Row Field
        
    init_rowcol field_name, xlRowField, position, _
            sort_by, sort_order, filter_type, filter_by_field_name, _
            filter_value1, filter_value2
    
End Sub


Sub init_col(field_name As String, _
        position As Long, _
        Optional sort_by As String, _
        Optional sort_order As XlSortOrder, _
        Optional filter_type As XlPivotFilterType, _
        Optional filter_by_field_name As String, _
        Optional filter_value1 As String, _
        Optional filter_value2 As String)
    ' Initialize Column Field
    
    init_rowcol field_name, xlColumnField, position, _
            sort_by, sort_order, filter_type, filter_by_field_name, _
            filter_value1, filter_value2
    
End Sub


Sub init_data(field_name As String, _
        source_field_name As String, _
        position As Long, _
        Optional data_function As XlConsolidationFunction, _
        Optional data_calculation As XlPivotFieldCalculation, _
        Optional data_number_format As String)
    ' Initialize Data Field
    
    init field_name, source_field_name, xlDataField, position
    
    If data_function = 0 Then
        p_data_function = xlSum
    Else
        p_data_function = data_function
    End If
    
    If data_calculation = 0 Then
        p_data_calculation = xlNoAdditionalCalculation
    Else
        p_data_calculation = data_calculation
    End If
    
    If data_number_format = "" Then
        p_data_number_format = "General"
    Else
        p_data_number_format = data_number_format
    End If

End Sub


Sub init_page(field_name As String, _
        position As Long, _
        Optional current_page As String)
    ' Initialize Page Field
    
    init field_name, field_name, xlPageField, position
    
    If current_page = "" Then
        p_current_page = "(All)"
    Else
        p_current_page = current_page
    End If
    
End Sub


Private Sub init_rowcol(field_name As String, _
        orientation As XlPivotFieldOrientation, _
        position As Long, _
        Optional sort_by As String, _
        Optional sort_order As XlSortOrder, _
        Optional filter_type As XlPivotFilterType, _
        Optional filter_by_field_name As String, _
        Optional filter_value1 As String, _
        Optional filter_value2 As String)
    ' Shared code for initializing Row and Column Fields
        
    init field_name, field_name, orientation, position
    
    p_rowcol_sort_by = sort_by
    
    If sort_order = 0 Then
        p_rowcol_sort_order = 1
    Else
        p_rowcol_sort_order = sort_order
    End If
    
    p_rowcol_filter_type = filter_type
    p_rowcol_filter_by_field_name = filter_by_field_name
    p_rowcol_filter_value1 = filter_value1
    p_rowcol_filter_value2 = filter_value2
    
End Sub


Private Sub init(field_name As String, _
        source_field_name As String, _
        orientation As XlPivotFieldOrientation, _
        position As Long)
    ' Shared code for initializing all 4 field orientations
    
    p_field_name = field_name
    p_source_field_name = source_field_name
    p_orientation = orientation
    p_position = position
    
End Sub


Public Function get_orientation() As XlPivotFieldOrientation
    get_orientation = p_orientation
End Function


Public Function set_orientation(orientation As XlPivotFieldOrientation) As Boolean
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


Public Function get_source_field_name() As String
    get_source_field_name = p_source_field_name
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
    get_rowcol_sort_by = p_rowcol_sort_by
End Function


Public Function set_rowcol_sort_by_field_name(Optional field_name As String) As Boolean

    ' if sort_by is vbnullstring then set sort_by to source_field_name
    If field_name = vbNullString Then
        p_rowcol_sort_by = p_source_field_name
    Else
        p_rowcol_sort_by = field_name
    End If
    
    set_rowcol_sort_by = True
    
End Function


Public Function set_rowcol_sort(sort_order As XlSortOrder, Optional sort_by As String) As Boolean
    set_rowcol_sort = set_rowcol_sort_by(sort_by) And set_rowcol_sort_order(sort_order)
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


Public Function get_current_page() As String
    get_current_page = p_current_page
End Function


Public Function set_current_page(current_page As String) As Boolean
    p_current_page = current_page
    set_current_page = True
End Function

