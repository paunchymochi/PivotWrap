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
    
    init_rowcol field_name, xlColumnField, position, _
            sort_by, sort_order, filter_type, filter_by_field_name, _
            filter_value1, filter_value2
    
End Sub


Sub init_data(field_name As String, _
        position As Long, _
        Optional data_function As XlConsolidationFunction, _
        Optional data_calculation As XlPivotFieldCalculation, _
        Optional data_number_format As String)
    
    init field_name, xlDataField, posiition
    
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
    
    init field_name, xlPageField, posiition
    
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
        
    init field_name, orientation, position
    
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
        orientation As XlPivotFieldOrientation, _
        position As Long)
    
    p_field_name = field_name
    p_orientation = orientation
    p_position = position
    
End Sub


Public Property Get orientation() As XlPivotFieldOrientation
    orientation = p_orientation
End Property

