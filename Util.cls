VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Util"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Public Function ValorExtenso(pdblValor As Double) As String
   Dim intQtde As Integer
   ValorExtenso = ""
   
      
End Function

Private Function RetornaExtenso() As String

   Select Case Numero
      Case 1: RetornaExtenso = "um"
      Case 2: RetornaExtenso = "dois"
      Case 3: RetornaExtenso = "três"
      Case 4: RetornaExtenso = "quatro"
      Case 5: RetornaExtenso = "cinco"
      Case 6: RetornaExtenso = "seis"
      Case 7: RetornaExtenso = "sete"
      Case 8: RetornaExtenso = "oito"
      Case 9: RetornaExtenso = "nove"
   End Select

End Function
