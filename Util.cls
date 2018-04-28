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

Public Function ValorExtenso(strValor As String) As String
    
   Dim Valor As Variant
   
   Valor = Split(strValor, ".")
   
   If UBound(Valor) > 1 Then
      'ca
   Else
   
   End If

   
   
     
        
End Function

Private Function RetornaExtensoCentavos(Numero As Integer) As String

   Select Case Numero
      Case 1: RetornaExtensoCentavos = "um"
      Case 2: RetornaExtensoCentavos = "dois"
      Case 3: RetornaExtensoCentavos = "três"
      Case 4: RetornaExtensoCentavos = "quatro"
      Case 5: RetornaExtensoCentavos = "cinco"
      Case 6: RetornaExtensoCentavos = "seis"
      Case 7: RetornaExtensoCentavos = "sete"
      Case 8: RetornaExtensoCentavos = "oito"
      Case 9: RetornaExtensoCentavos = "nove"
   End Select
   
   
End Function

Private Function RetornaExtensoDezenas(Numero As Integer) As String

   Select Case Numero
      Case 2: RetornaExtensoDezenas = "vinte"
      Case 3: RetornaExtensoDezenas = "trinta"
      Case 4: RetornaExtensoDezenas = "quarenta"
      Case 5: RetornaExtensoDezenas = "cinquenta"
      Case 6: RetornaExtensoDezenas = "sessenta"
      Case 7: RetornaExtensoDezenas = "setenta"
      Case 8: RetornaExtensoDezenas = "oitenta"
      Case 9: RetornaExtensoDezenas = "noventa"
   End Select

End Function


Private Function RetornaExtensoCentenas(Numero As Integer) As String

   Select Case Numero
      Case 1: RetornaExtensoCentenas = "cem"
      Case 2: RetornaExtensoCentenas = "duzentos"
      Case 3: RetornaExtensoCentenas = "trezentos"
      Case 4: RetornaExtensoCentenas = "quatrocentos"
      Case 5: RetornaExtensoCentenas = "quinhentos"
      Case 6: RetornaExtensoCentenas = "seiscentos"
      Case 7: RetornaExtensoCentenas = "setecentos"
      Case 8: RetornaExtensoCentenas = "oitocentos"
      Case 9: RetornaExtensoCentenas = "novecentos"
   End Select

End Function

Private Function RetornaExtensoMilhares(Numero As Integer) As String

   Select Case Numero
      Case 1: RetornaExtensoMilhares = "mil"
      Case 2: RetornaExtensoMilhares = "duzzento"
      Case 3: RetornaExtensoMilhares = "três"
      Case 4: RetornaExtensoMilhares = "quatro"
      Case 5: RetornaExtensoMilhares = "cinco"
      Case 6: RetornaExtensoMilhares = "seis"
      Case 7: RetornaExtensoMilhares = "sete"
      Case 8: RetornaExtensoMilhares = "oito"
      Case 9: RetornaExtensoMilhares = "nove"
   End Select

End Function


