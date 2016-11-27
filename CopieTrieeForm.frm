VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} CopieTrieeForm 
   Caption         =   "Copie et trie une plage"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   2310
   OleObjectBlob   =   "CopieTrieeForm.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "CopieTrieeForm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
  Dim PlageCellules As Range
  Dim NColonneTriee As Long
  Dim n As Integer
  Set PlageCellules = Range(SelectionneurPlage.Value)
  NColonneTriee = CLng(SaisieNColonne.Value)
  ' Les sous-procedures s'invoquent sans les parentheses
  CopieTriee Plage:=PlageCellules, NColonneTriee:=NColonneTriee
  
End Sub



' Vous pouvez ecrire une instruction sur plusieurs lignes, a condition
' de bien preceder chaque retour a la ligne d'un "_" .

Public Sub CopieTriee(ByVal Plage As Range, _
 NColonneTriee As Long, Optional IgnoreEnTete As XlYesNoGuess = xlNo)

  Dim Lignes As Integer
  Lignes = Plage.Rows.Count
  Dim NouvelleFeuille As Worksheet
  Dim PlageCopie As Range
  Set NouvelleFeuille = ThisWorkbook.Worksheets.Add
  Set PlageCopie = _
    NouvelleFeuille.Range("A1").Resize(Lignes, Plage.Columns.Count)
  Plage.Copy PlageCopie
  
  Dim ColonneTriee As Range

  Set ColonneTriee = PlageCopie.Range( _
    PlageCopie.Cells(1, NColonneTriee), _
    PlageCopie.Cells(Lignes, NColonneTriee) _
  )

  ' La syntaxe nom_param:=valeur permet de passer une valeur a une
  ' fonction, en specifiant le nom du parametre regle.
  '
  ' Pour des fonctions ayant enormements de parametres, comme Sort,
  ' ce genre de methodologie est fort conseillee.
  ' Sort (Anglais) <-> Tri (Francais)

  PlageCopie.Sort key1:=ColonneTriee, order1:=xlAscending, _
                  Header:=IgnoreEnTete
  
End Sub


