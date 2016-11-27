Attribute VB_Name = "Module1"
' Code d'essai

' "Passer" des arguments a une fonction signifie "copier" les valeurs
' passees en tant qu'arguments.
'
' Par exemple, prenons une fonction definie comme ceci :
'
' Function Additioner(FirstOp As Integer, SecondOp As Integer)
'   Additioner = FirstOp + SecondOp
' End Function
'
' Lorsque l'on invoque Additioner(1, 2) de cette maniere les valeurs
' "1" et "2" sont "passees en argument".
'
' Cependant, il est necessaire de comprendre que la copie peut se faire
' de deux moyens differents.
'
'  ByVal indique que toutes les donnees contenues dans la variable sont
' elles meme sont copiees en memoire. La copie etant accessible via
' l'argument receveur. C'est le "passage par valeur".
'
'  A l'inverse ByRef indique que la reference vers ces donnees est
' copiee. Une "reference" etant globalement l'adresse memoire ou se
' trouvent les donnees. C'est le "passage par reference"
'
' L'avantage des references est qu'elles ne consomment quasiment pas
' de memoire. (4 voire 8 octets generalement)
' L'inconvenient est que toute modification effectuee sur une reference
' modifie l'element qui est reference, par definition.
'
' Inversement, passer les valeurs par copie consomme plus de memoire
' mais evite de modifier les donnees originales.
'
' En cas de modification durant la fonction, mieux vaut que les valeurs
' soient passees par copie, quitte a perdre un peu de memoire. Les
' references sont utiles si vous ne faites que lire les donnees, ou
' bien lorsque vous avez a traiter de TRES gros volumes de donnees
' (Plusieurs centaines de Mo de donnees), ou lorsque vous pouvez vous
' permettre de modifier les valeurs sans problemes.
'
' Note : La reaffectation d'un parametre contenant une reference
'        modifie le contenu de valeur passee en argument.
'        C'est a dire que dans l'exemple suivant, SecondOp se voit
'        reaffecter la valeur '75'. Comme SecondOp est un parametre
'        contenant une reference, la reaffectation va modifier la
'        valeur de 'b' dans TestAdditioner !
' ---

Function BadAdd(ByVal FirstOp As Integer, ByRef SecondOp As Integer)

  FirstOp = 10
  SecondOp = 75
  BadAdd = SecondOp
End Function

Sub TestBadAdd()
  Dim a As Integer
  Dim b As Integer
  Dim result As Integer
  a = 1
  b = 2
  
  MsgBox ("Avant : a : " & a & " - b : " & b)
  
  result = BadAdd(a, b)
  ' a -> FirstOp (ByVal)
  ' b -> SecondOp (ByRef)
  ' La valeur 'a' est copiee
  ' La reference de 'b' est copiee
  ' La reaffection de FirstOp dans BadAdd n'affectera pas 'a' car
  ' 'a' est copie.
  ' La reaffectation de SecondOp dans BadAdd affectera 'b' car 'b'
  ' et SecondOp font reference a la meme valeur !
  
  MsgBox ("Apres : a : " & a & " - b : " & b)
  
End Sub
