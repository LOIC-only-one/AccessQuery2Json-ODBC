## 🛠️ Description  
Ce script en VBScript permet d'exécuter une requête SQL via une connexion ODBC, de formater les résultats en JSON, et de les écrire dans un fichier. Il est utile pour automatiser des tâches d'extraction de données depuis une base de données Access ou autre source compatible ODBC.  

---

## 📋 Fonctionnalités  
- 🔒 **EchapperChaineJson** : Échappe les caractères spéciaux pour le format JSON.  
- 🛠️ **ExecuterRequete** : Exécute une requête SQL via ODBC et retourne les résultats au format JSON.  
- 📝 **EcrireDansFichier** : Écrit le contenu JSON dans un fichier.  
- 🚀 **main** : Point d'entrée du script, qui orchestre les différentes étapes.  

---

## ⚙️ Configuration  

### 1️⃣ **Configurer ODBC**  
1. Ouvrez le **Gestionnaire de sources de données ODBC** (odbcad32.exe).  
2. Allez dans l'onglet **DSN Système** ou **DSN Utilisateur**.  
3. Cliquez sur **Ajouter** et sélectionnez le driver **Microsoft Access Driver (*.mdb, *.accdb)**.  
4. Donnez un nom au DSN (par exemple : `verif`).  
5. Sélectionnez le fichier Access (.mdb ou .accdb) à utiliser comme source de données.  

### 2️⃣ **Vérifier le driver ODBC**  
- Assurez-vous que le driver ODBC pour Access est installé.  
- Pour les systèmes 64 bits, utilisez le gestionnaire ODBC 32 bits situé dans `C:\Windows\SysWOW64\odbcad32.exe`.  

---

## 🖥️ Utilisation  

1. **Modifier les paramètres** :  
    - `nomDsn` : Nom du DSN configuré (par défaut : `verif`).  
    - `requete` : Requête SQL à exécuter (par défaut : `SELECT * FROM VIP`).  
    - `cheminFichierSortie` : Chemin du fichier de sortie JSON.  

2. **Exécuter le script** :  
    - Sauvegardez le script avec l'extension `.vbs`.  
    - Double-cliquez sur le fichier ou exécutez-le via la ligne de commande :  
      ```bash
      cscript script.vbs
      ```  

---

## 📂 Exemple de sortie  
Un fichier JSON sera généré avec le contenu suivant :  
```json
[
  {
     "Champ1": "Valeur1",
     "Champ2": "Valeur2"
  },
  {
     "Champ1": "Valeur3",
     "Champ2": "Valeur4"
  }
]
```  

---

## 🛑 Limitations  
- Le script ne gère pas les bases de données nécessitant une authentification complexe.  
- Les performances peuvent être limitées pour des bases de données volumineuses.  

---

## 📌 Notes  
- Assurez-vous que le fichier Access n'est pas verrouillé par un autre processus.  
- Vérifiez les permissions d'écriture sur le chemin de sortie.  

---  
🎉 **Amusez-vous bien avec ce script VBScript !**  
