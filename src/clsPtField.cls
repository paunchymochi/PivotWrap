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
Private p_field_name_original As String
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
        Position As Long, _
        Optional sort_by As String, _
        Optional sort_order As XlSortOrder, _
        Optional filter_type As XlPivotFilterType, _
        Optional filter_by_field_name As String, _
        Optional filter_value1 As String, _
        Optional filter_value2 As String)
    ' Initialize Row Field
        
    init_rowcol field_name, xlRowField, Position, _
            sort_by, sort_order, filter_type, filter_by_field_name, _
            filter_value1, filter_value2
    
End Sub


Sub init_col(field_name As String, _
        Position As Long, _
        Optional sort_by As String, _
        Optional sort_order As XlSortOrder, _
        Optional filter_type As XlPivotFilterType, _
        Optional filter_by_field_name As String, _
        Optional filter_value1 As String, _
        Optional filter_value2 As String)
    ' Initialize Column Field
    
    init_rowcol field_name, xlColumnField, Position, _
            sort_by, sort_order, filter_type, filter_by_field_name, _
            filter_value1, filter_value2
    
End Sub


Sub init_data(field_name As String, _
        field_name_original As String, _
        Position As Long, _
        Optional data_function As XlConsolidationFunction, _
        Optional data_calculation As XlPivotFieldCalculation, _
        Optional data_number_format As String)
    ' Initialize Data Field
    
    init field_name, field_name_original, xlDataField, Position
    
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
        Position As Long, _
        Optional current_page As String)
    ' Initialize Page Field
    
    init field_name, field_name, xlPageField, Position
    
    If current_page = "" Then
        p_current_page = "(All)"
    Else
        p_current_page = current_page
    End If
    
End Sub


Private Sub init_rowcol(field_name As String, _
        Orientation As XlPivotFieldOrientation, _
        Position As Long, _
        Optional sort_by As String, _
        Optional sort_order As XlSortOrder, _
        Optional filter_type As XlPivotFilterType, _
        Optional filter_by_field_name As String, _
        Optional filter_value1 As String, _
        Optional filter_value2 As String)
    ' Shared code for initializing Row and Column Fields
        
    init field_name, field_name, Orientation, Position
    
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
        field_name_original As String, _
        Orientation As XlPivotFieldOrientation, _
        Position As Long)
    ' Shared code for initializing all 4 field orientations
    
    p_field_name = field_name
    p_field_name_original = field_name_original
    p_orientation = Orientation
    p_position = Position
    
End Sub

Function get_orientation() As XlPivotFieldOrientation
    get_orientation = p_orientation
End Function


Function get_field_name() As String
    get_field_name = p_field_name
End Function


Function get_field_name_original() As String
    get_field_name_original = p_field_name_original
End Function


Function get_position() As Long
    get_position = p_position
End Function


Function get_rowcol_sort_order() As XlSortOrder
    get_rowcol_sort_order = p_rowcol_sort_order
End Function


Function get_rowcol_sort_by() As String
    get_rowcol_sort_by = p_rowcol_sort_by
End Function


Function get_rowcol_filter_type() As XlPivotFilterType
    get_rowcol_filter_type = p_rowcol_filter_type
End Function


Function get_rowcol_filter_by_field_name() As String
    get_rowcol_filter_by_field_name = p_rowcol_filter_by_field_name
End Function


Function get_rowcol_filter_value1() As String
    get_rowcol_filter_value1 = p_rowcol_filter_value1
End Function


Function get_rowcol_filter_value2() As String
    get_rowcol_filter_value2 = p_rowcol_filter_value2
End Function


Function get_data_function() As XlConsolidationFunction
    get_data_function = p_data_function
End Function


Function get_data_calculation() As XlPivotFieldCalculation
    get_data_calculation = p_data_calculation
End Function


Function get_data_number_format() As String
    get_data_number_format = p_data_number_format
End Function


Function get_repeat_labels() As XlPivotFieldRepeatLabels
    get_repeat_labels = p_repeat_labels
End Function


Function get_current_page() As String
    get_current_page = p_current_page
End Function

