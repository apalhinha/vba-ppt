VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "cStack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
' Shared Powerpoint / Excel

Dim pStack As Collection

Public Function Pop() As Variant
    With pStack
        If .Count > 0 Then
            If IsObject(.Item(.Count)) Then
                Set Pop = .Item(.Count)
            Else
                Pop = .Item(.Count)
            End If
            .Remove .Count
        End If
    End With
End Function

Public Function Push(newItem As Variant) As Variant
    With pStack
        .Add newItem
        If IsObject(newItem) Then
            Set Push = .Item(.Count)
        Else
            Push = .Item(.Count)
        End If
    End With

End Function

Public Sub Init()
    Set pStack = New Collection
End Sub

Public Property Get IsEmpty() As Boolean
    IsEmpty = (pStack.Count = 0)
End Property

