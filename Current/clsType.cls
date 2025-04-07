Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Compare Database
Option Explicit

Private lngID As Long
Private strName As String

Public Property Get ID() As Long

    ID = lngID


End Property
Public Property Let ID(lngID1 As Long)

    lngID = lngID1

End Property
Public Property Get Name() As String

    Name = strName


End Property
Public Property Let Name(strName1 As String)

    strName = strName1

End Property