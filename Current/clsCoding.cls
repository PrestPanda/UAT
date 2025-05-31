Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Compare Database
Option Explicit

Dim Log As New clsLog


Public Enum enuCoding_Visibility

    vPrivate
    vPublic

End Enum
Public Enum enuCoding_DataTypes

    vtNoType
    vtString
    vtStringLong
    vtLong
    vtDouble
    vtBoolean
    vtDate
    vtVariant
    vtObject
    vtAuto

End Enum
Public Enum enuCoding_Method_Types

    mtSub
    mtFunction

End Enum

Public Function Get_Class_Reference_DimensionStatement(ByVal strName As String) As String

    Get_Class_Reference_DimensionStatement = "Dim " & _
        Replace(strName, "cls", "") & " as new " & strName

End Function
Public Function Get_Variable_DimensionStatement_ByEnum(Visibility As enuCoding_Visibility, strName As String, _
    DataType As enuCoding_DataTypes) As String
'Refactored
'Returns a dimension statement for a variable based on visability and data type

    Get_Variable_DimensionStatement_ByEnum = Get_Visibility_AsString_ByEnum(Visibility) & " " & _
                                        Get_Variable_Name_AsString_ByEnum(strName, DataType) & " as " & _
                                        Get_Variable_Type_AsString_ByEnum(DataType)

End Function
Public Function Get_Variable_DimensionStatement_ByString(strVisibility As String, strName As String, _
    strDataType As String) As String
'Refactored
'Returns a dimension statement for a variable based on visability and data type

    Get_Variable_DimensionStatement_ByString = strVisibility & " " & _
                                        Get_Variable_Name_AsString_ByString(strName, strDataType) & " as " & _
                                        Get_Variable_Translate_Datatype(strDataType)

End Function
Public Function Get_Variable_Name_AsString_ByEnum( _
    strName As String, _
    DataType As enuCoding_DataTypes) As String
'Refactored
'Returns the name of a Variable with variable type prefix

    Get_Variable_Name_AsString_ByEnum = Get_Variable_Prefix_AsString_ByEnum(DataType) & strName

End Function
Public Function Get_Variable_Name_AsString_ByString( _
    ByVal strName As String, _
    ByVal strDataType As String) As String
'Refactored
'Returns the name of a Variable with variable type prefix

    Get_Variable_Name_AsString_ByString = Get_Variable_Prefix_AsString_ByString(strDataType) & strName

End Function
Public Function Get_Visibility_AsString_ByEnum(Visibilty As enuCoding_Visibility) As String
'Refactored
'Returns the visability type as string

    Select Case Visibilty
    
        Case vPrivate
            Get_Visibility_AsString_ByEnum = "Private"
            
        Case vPublic
            Get_Visibility_AsString_ByEnum = "Public"
            
    End Select

End Function
Public Function Get_Visibility_AsEnum_ByString(strVisibilty As String) As enuCoding_Visibility
'Refactored
'Returns the visability type as enumeration

    Select Case strVisibilty
    
        Case "Private"
            Get_Visibility_AsEnum_ByString = vPrivate
            
        Case "Public"
            Get_Visibility_AsEnum_ByString = vPublic
            
    End Select

End Function
Public Function Get_Variable_Prefix_AsString_ByString(strDataType As String) As String
'Refactored

    Get_Variable_Prefix_AsString_ByString = _
        Get_Variable_Prefix_AsString_ByEnum( _
        Get_Variable_Type_AsEnum_ByString(strDataType))

End Function

Public Function Get_Variable_Prefix_AsString_ByEnum(VariableType As enuCoding_DataTypes) As String
'Refactored
'Returns the prefix of a variable with a specific data type

    Select Case VariableType
    
        Case vtAuto
            Get_Variable_Prefix_AsString_ByEnum = ""
    
        Case vtBoolean
            Get_Variable_Prefix_AsString_ByEnum = "bln"
        
        Case vtDate
            Get_Variable_Prefix_AsString_ByEnum = "dat"
            
        Case vtDouble
            Get_Variable_Prefix_AsString_ByEnum = "dbl"
            
        Case vtLong
            Get_Variable_Prefix_AsString_ByEnum = "lng"
            
        Case vtString, vtStringLong
            Get_Variable_Prefix_AsString_ByEnum = "str"
            
        Case vtVariant
            Get_Variable_Prefix_AsString_ByEnum = ""
            
        Case Else
            Get_Variable_Prefix_AsString_ByEnum = ""
            
    End Select
    

End Function
Public Function Get_Variable_Type_AsString_ByEnum(VariableType As enuCoding_DataTypes) As String
'Refactored
'Returns a string with the name of the Data Type

    Select Case VariableType
    
        Case vtBoolean
            Get_Variable_Type_AsString_ByEnum = "Boolean"
        
        Case vtDate
            Get_Variable_Type_AsString_ByEnum = "Date"
            
        Case vtDouble
            Get_Variable_Type_AsString_ByEnum = "Double"
            
        Case vtLong
            Get_Variable_Type_AsString_ByEnum = "Long"
            
        Case vtString, vtStringLong
            Get_Variable_Type_AsString_ByEnum = "String"
            
        Case vtVariant
            Get_Variable_Type_AsString_ByEnum = "Variant"
            
        Case vtNoType
            Get_Variable_Type_AsString_ByEnum = ""
    
    End Select

End Function
Public Function Get_Variable_Translate_Datatype(strDataType As String) As String

    Select Case strDataType
        
        Case "Long String"
            Get_Variable_Translate_Datatype = "String"
            
        Case Else
            Get_Variable_Translate_Datatype = strDataType
            
    End Select

End Function

Public Function Get_Variable_Type_AsEnum_ByString(ByVal strDataType As String) As enuCoding_DataTypes
'Refactored
'Returns

    Select Case strDataType
    
        Case "Boolean"
            Get_Variable_Type_AsEnum_ByString = vtBoolean
        
        Case "Date"
            Get_Variable_Type_AsEnum_ByString = vtDate
            
        Case "Double"
            Get_Variable_Type_AsEnum_ByString = vtDouble
            
        Case "Long"
            Get_Variable_Type_AsEnum_ByString = vtLong
            
        Case "String", "Long String"
            Get_Variable_Type_AsEnum_ByString = vtString
            
        Case "Variant"
            Get_Variable_Type_AsEnum_ByString = vtVariant
            
        Case Else
            Get_Variable_Type_AsEnum_ByString = vtNoType
    
    End Select

End Function
Public Function Get_Method_Type_AsString_ByEnum(MethodType As enuCoding_Method_Types) As String
'Refactored
'Returns the type of a Method (Sub/Function)

    Select Case MethodType
    
        Case mtSub
            Get_Method_Type_AsString_ByEnum = "Sub"
            
        Case mtFunction
            Get_Method_Type_AsString_ByEnum = "Function"
    
    End Select

End Function
Public Function Get_Method_Type_AsEnum_ByString(strMethodType As String) As enuCoding_Method_Types
'Refactored
'Returns the type of a Method (Sub/Function)

    Select Case strMethodType
    
        Case "Sub"
            Get_Method_Type_AsEnum_ByString = mtSub
            
        Case "Function"
            Get_Method_Type_AsEnum_ByString = mtFunction
    
    End Select

End Function
    