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
Public Enum enuCoding_Variable_Types

    vtNoType
    vtString
    vtStringLong
    vtlong
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


Public Sub ClassModule_CreateNew(strClassName As String, Optional strVBACode As String)
'Creates a new class module and inserts VBA Code

    Dim objVBComp As New VBComponent


    Set objVBComp = Application.VBE.ActiveVBProject.VBComponents.Add(vbext_ct_ClassModule)
    objVBComp.CodeModule.AddFromString strVBACode


    Log.WriteLine "Bitte benenne die Klasse '" & objVBComp.Name & "' in '" & strClassName & "' um."


End Sub
Public Function Get_Method_Type(MethodType As enuCoding_Method_Types) As String
'Returns the type of a Method (Sub/Function)

    Select Case MethodType
    
        Case mtSub
            Get_Method_Type = "Sub"
            
        Case mtFunction
            Get_Method_Type = "Function"
    
    End Select

End Function
Public Function Get_Variable_DimensionStatement(Visibility As enuCoding_Visibility, strName As String, _
    VariableType As enuCoding_Variable_Types) As String
'Returns a dimension statement for a variable based on visability and data type

    Get_Variable_DimensionStatement = Get_Visibility(Visibility) & " " & _
                                        Get_Variable_Name(strName, VariableType) & " as " & _
                                        Get_Variable_Type(VariableType)

End Function
Public Function Get_Variable_Name(strPropertyName As String, VariableType As enuCoding_Variable_Types) As String
'Returns the name of a Variable with variable type prefix

    Get_Variable_Name = Get_Variable_Prefix(VariableType) & strPropertyName

End Function
Public Function Get_Variable_Prefix(VariableType As enuCoding_Variable_Types) As String
'Returns the prefix of a variable with a specific data type

    Select Case VariableType
    
        Case vtBoolean
            Get_Variable_Prefix = "bln"
        
        Case vtDate
            Get_Variable_Prefix = "dat"
            
        Case vtDouble
            Get_Variable_Prefix = "dbl"
            
        Case vtlong, vtAuto
            Get_Variable_Prefix = "lng"
            
        Case vtString, vtStringLong
            Get_Variable_Prefix = "str"
            
        Case vtVariant
            Get_Variable_Prefix = "var"
            
        Case vtNoType
            Get_Variable_Prefix = ""
            
    End Select
    

End Function
Public Function Get_Variable_Type(VariableType As enuCoding_Variable_Types) As String
'Returns a string with the name of the Data Type
    Select Case VariableType
    
        Case vtBoolean
            Get_Variable_Type = "Boolean"
        
        Case vtDate
            Get_Variable_Type = "Date"
            
        Case vtDouble
            Get_Variable_Type = "Double"
            
        Case vtlong, vtAuto
            Get_Variable_Type = "Long"
            
        Case vtString, vtStringLong
            Get_Variable_Type = "String"
            
        Case vtStringLong
            Get_Variable_Type = "Long"
            
        Case vtVariant
            Get_Variable_Type = "Variant"
            
        Case vtNoType
            Get_Variable_Type = ""
    
    End Select

End Function
Public Function Get_Visibility(Visibilty As enuCoding_Visibility) As String
'Returns the visability type

    Select Case Visibilty
    
        Case vPrivate
            Get_Visibility = "Private"
            
        Case vPublic
            Get_Visibility = "Public"
    End Select

End Function