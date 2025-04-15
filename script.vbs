Option Explicit

Function EchapperChaineJson(valeur)
    valeur = Replace(valeur, "\", "\\")
    valeur = Replace(valeur, """", "\""")
    valeur = Replace(valeur, vbCrLf, "\n")
    valeur = Replace(valeur, vbTab, "\t")
    EchapperChaineJson = valeur
End Function

Function ExecuterRequete(nomDsn, requeteSql)
    On Error Resume Next
    
    Dim connexion, resultat, nombreChamps, nombreLignes
    Dim i, nomChamp, valeurChamp, messageErreur, resultatFinal, tableauJson, objetJson
    
    Set connexion = CreateObject("ADODB.Connection")
    connexion.Open "DSN=" & nomDsn
    
    If Err.Number <> 0 Then
        messageErreur = "Erreur de connexion ODBC: " & Err.Number & " - " & Err.Description
        WScript.Echo messageErreur
        ExecuterRequete = "Erreur: " & messageErreur
        Exit Function
    End If
    
    Set resultat = CreateObject("ADODB.Recordset")
    resultat.Open requeteSql, connexion, 3, 1 
    
    If resultat.EOF And resultat.BOF Then
        resultatFinal = "[]"
    Else
        nombreChamps = resultat.Fields.Count
        tableauJson = "["
        
        Do Until resultat.EOF
            objetJson = "{"
            For i = 0 To nombreChamps - 1
                nomChamp = resultat.Fields(i).Name
                valeurChamp = resultat.Fields(i).Value
                
                If IsNull(valeurChamp) Then
                    valeurChamp = "null"
                Else
                    valeurChamp = """" & EchapperChaineJson(CStr(valeurChamp)) & """"
                End If
                
                objetJson = objetJson & """" & nomChamp & """: " & valeurChamp
                If i < nombreChamps - 1 Then
                    objetJson = objetJson & ", "
                End If
            Next
            objetJson = objetJson & "}"
            tableauJson = tableauJson & objetJson
            resultat.MoveNext
            If Not resultat.EOF Then
                tableauJson = tableauJson & ", "
            End If
        Loop
        tableauJson = tableauJson & "]"
        resultatFinal = tableauJson
    End If

    resultat.Close
    connexion.Close
    Set resultat = Nothing
    Set connexion = Nothing
    
    ExecuterRequete = resultatFinal
End Function

Sub EcrireDansFichier(cheminFichier, contenu)
    Dim fso, fichier, flux
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set flux = CreateObject("ADODB.Stream")
    
    flux.Type = 2
    flux.Charset = "utf-8"
    flux.Open
    flux.WriteText contenu
    flux.SaveToFile cheminFichier, 2
    flux.Close
    
    Set flux = Nothing
    Set fso = Nothing
End Sub

Sub main()
    Dim nomDsn, requete, resultat, cheminFichierSortie
    nomDsn = "verif"
    requete = "SELECT * FROM VIP"
    cheminFichierSortie = "C:\...\output.json"
    
    resultat = ExecuterRequete(nomDsn, requete)
    WScript.Echo resultat
    
    EcrireDansFichier cheminFichierSortie, resultat
    WScript.Echo "Resultat ecrit dans le fichier : " & cheminFichierSortie
End Sub
main
