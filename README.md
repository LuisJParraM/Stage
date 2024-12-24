# Automated Form Generation Project Documentation

This repository documents the development and functionality of a Microsoft Access-based tool to automate shipment form generation. The tool uses VBA and SQL for seamless data handling and user interaction. Below are detailed sections explaining the process, functionality, and code implementation.

## Project Overview

The project aims to:

- Automate the extraction and integration of data from 146 manually created Excel forms.

- Provide technicians with an intuitive interface for managing internal requesters, recipients, and references.

- Ensure data accuracy through automated cleanup processes.

- Enable easy export of shipment forms to preformatted Excel templates.

## Main Components:

1. **Data Import:** Automating data extraction from Excel into Access tables.
2. **Data Cleanup:** SQL-based deduplication for maintaining data integrity.
3. **User Forms:**
   - Form 1 (Formulaire 1) : Acts as a navigation menu.
   - Form 2 (Formulaire 2) : Provides detailed functionalities for data input and export.

## Data Import and Cleanup

### Data Import (VBA Code):

The following VBA code extracts data from Excel files and imports it into the Access database:

### Código VBA: Extract_donne_Click

```vbnet
Private Sub Extract_donne_Click()
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim rutaCarpeta As String
    Dim rutaArchivo As String
    Dim archivo As String
    Dim xlApp As Object
    Dim xlBook As Object
    Dim xlSheet As Object

    ' Folder path and Access database path
    rutaCarpeta = "K:\A_27_TS\A_2704_TSM\A_270402_Maintenance\Formulaire d'expédition"
    Dim rutaDB As String
    rutaDB = "K:\A_27_TS\A_2704_TSM\A_270402_Maintenance\Formulaire d'expédition\Formulaire d'expédition.accdb"

    ' Open Access database connection
    Set db = DBEngine.OpenDatabase(rutaDB)

    ' Iterate through Excel files in the folder
    archivo = Dir(rutaCarpeta & "\*.xls*")
    Do While archivo <> ""
        rutaArchivo = rutaCarpeta & "\" & archivo

        ' Open Excel file
        Set xlApp = CreateObject("Excel.Application")
        Set xlBook = xlApp.Workbooks.Open(rutaArchivo, ReadOnly:=True)
        Set xlSheet = xlBook.Sheets(1)

        ' Read cell values
        Dim valores(1 To 3) As Variant
        valores(1) = xlSheet.Range("C13").Value ' Company name
        valores(2) = xlSheet.Range("C23").Value ' Merchandise reference
        valores(3) = xlSheet.Range("C24").Value ' Merchandise designation

        ' Insert values into Access table
        Set rs = db.OpenRecordset("Reference", dbOpenDynaset)
        rs.AddNew
        rs![destinataire] = valores(1)
        rs![Référance] = valores(2)
        rs![Désignation] = valores(3)
        rs.Update
        rs.Close

        ' Close Excel file
        xlBook.Close False
        xlApp.Quit

        ' Release objects
        Set xlSheet = Nothing
        Set xlBook = Nothing
        Set xlApp = Nothing

        ' Next file
        archivo = Dir
    Loop

    ' Close database connection
    db.Close
    Set db = Nothing

    MsgBox "Data imported successfully into the Access database.", vbInformation
End Sub
```
**Purpose:** This code automates the extraction of company names, merchandise references, and designations from Excel sheets and stores them in the Reference table in Access.

### Data Cleanup (SQL Queries): 

#### Identifying Duplicates:

```sql
SELECT [Destinataire], [Référance], [Désignation], COUNT(*) AS TotalDuplicados 
FROM Reference 
GROUP BY [Destinataire], [Référance], [Désignation] 
HAVING COUNT(*) > 1;
```
**Explanation:** This query identifies duplicate records in the Reference table based on recipient, reference, and designation.

#### Removing Duplicates:

```sql
DELETE FROM Reference 
WHERE [N°] NOT IN (
    SELECT MIN([N°]) 
    FROM Reference 
    GROUP BY [Nom de l'entreprise], [Nom du destinataire], [Adresse: rue], [Adresse: CP-Ville], [Adresse: Pays], [Numéro de téléphone], [Adresse mail]
);
```
**Explanation:** This query retains the first occurrence of duplicate records while removing others.

## Form 1 (Formulaire1): Navigation Menu

### Purpose: 
Formulaire1 serves as the entry point to the application, providing a simple interface for navigating to the detailed data input form or exiting the application.

### VBA Code:

```vbnet
Private Sub Bouton_creation_de_formulaire_dexpedition_Click()
    ' Opens Formulaire2 in normal view.
    DoCmd.OpenForm "Formulaire2", acNormal
End Sub

Private Sub Bouton_quitter_Click()
    Dim confirmation As Integer

    ' Ask for confirmation before exiting Access.
    confirmation = MsgBox("Voulez-vous vraiment quitter Access ?", vbYesNo + vbQuestion, "Confirmation")

    If confirmation = vbYes Then
        ' Close the current form and quit Access.
        DoCmd.Close acForm, Me.Name
        Application.Quit acQuitSaveAll
    End If
End Sub
```
#### Explanation:
  1. **Bouton_creation_de_formulaire_dexpedition_Click:** Opens Formulaire2, the main form where data is managed.
  2. **Bouton_quitter_Click:** Prompts the user for confirmation before closing the form and exiting the Access application.

## Form 2 (Formulaire2): Provides detailed functionalities for data input and export
Formulaire2 serves as the main data management interface for:
  - **Managing Internal Requesters:** Dynamically add and select requesters.

  - **Managing Recipients:** Dynamically add and select recipients.

  - **Handling References:** Link references to recipients and manage their details.

  - **Exporting Data to Excel:** Generate preformatted shipment forms for external use.

### 1. Managing Internal Requesters

#### Code: Handling Selection of Existing Requesters

```vbnet
Private Sub Modifiable_demandeur_interne_AfterUpdate()
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim idNom As Variant

    ' Capture the selected value
    idNom = Me.Modifiable_demandeur_interne.Value

    ' Validate that the value is not empty
    If Trim(idNom & "") = "" Then
        MsgBox "Veuillez sélectionner une valeur ou saisir un nouveau demandeur.", vbExclamation, "Erreur"
        Exit Sub
    End If

    ' Search for the selected value in the table
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT * FROM [Demandeur interne] WHERE [N°] = " & idNom, dbOpenSnapshot)

    If Not rs.EOF Then
        ' If the value exists, fill in the labels
        Me.étiquette_demandeur_interne_1.Caption = rs![nom]
        Me.étiquette_demandeur_interne_2.Caption = rs![Numéro de téléphone]
        Me.étiquette_demandeur_interne_3.Caption = rs![Adresse mail]
        Me.étiquette_demandeur_interne_4.Caption = rs![Service à imputer]
        Me.étiquette_demandeur_interne_5.Caption = 22240
        Me.étiquette_date_de_livraison_1.Caption = Date
    Else
        MsgBox "Erreur : La sélection n'existe pas dans la base de données.", vbCritical, "Erreur"
    End If

    rs.Close
    Set rs = Nothing
    Set db = Nothing
End Sub
```
**Explanation:**

  * This subroutine triggers when the user selects a requester from the dropdown.

  * It validates the input and fetches the corresponding requester details from the database.

  * The details are displayed in the form labels for user reference.

#### Code: Adding New Requesters

```vbnet
Private Sub Modifiable_demandeur_interne_NotInList(NewData As String, Response As Integer)
    Dim sqlInsert As String
    Dim userResponse As Integer
    Dim nom As String
    Dim Telephone As String
    Dim adresseMail As String
    Dim service As String

    ' Ask if the user wants to add the new value
    userResponse = MsgBox("Le demandeur '" & NewData & "' n'existe pas. Voulez-vous l'ajouter ?", vbYesNo + vbQuestion, "Nouveau Demandeur")

    If userResponse = vbYes Then
        ' Request additional information
        nom = InputBox("Veuillez saisir le nom :", "Nouveau Demandeur")
        Telephone = InputBox("Veuillez saisir le numéro de téléphone :", "Nouveau Demandeur")
        adresseMail = InputBox("Veuillez saisir l'adresse mail :", "Nouveau Demandeur")
        service = "Maintenance" ' Default value for "Service à imputer"

        ' Validate that the fields are not empty
        If Trim(nom) = "" Or Trim(Telephone) = "" Or Trim(adresseMail) = "" Then
            MsgBox "Tous les champs sont obligatoires. L'ajout a été annulé.", vbExclamation, "Erreur"
            Response = acDataErrContinue
            Exit Sub
        End If

        ' Insert the new record into the table
        sqlInsert = "INSERT INTO [Demandeur interne] ([nom], [Numéro de téléphone], [Adresse mail], [Service à imputer]) VALUES ('" & Replace(nom, "'", "''") & "', '" & Replace(Telephone, "'", "''") & "', '" & Replace(adresseMail, "'", "''") & "', '" & Replace(service, "'", "''") & "');"
        CurrentDb.Execute sqlInsert, dbFailOnError

        ' Update the dropdown list
        Me.Modifiable_demandeur_interne.Undo ' Cancel ongoing changes
        Me.Modifiable_demandeur_interne.RowSource = "SELECT * FROM [Demandeur interne];"
        Me.Modifiable_demandeur_interne.Requery

        ' Notify Access that the new data was successfully added
        Response = acDataErrAdded

        MsgBox "Le demandeur '" & NewData & "' a été ajouté avec succès.", vbInformation, "Succès"
    Else
        ' Cancel if the user does not want to add the value
        Response = acDataErrContinue
        MsgBox "Le demandeur n'a pas été ajouté.", vbExclamation, "Action Annulée"
    End If
End Sub
```
**Explanation:**

  * This subroutine triggers when a user enters a value not present in the dropdown list.

  * It prompts the user to confirm the addition and collect further details.

  * A new record is added to the Demandeur interne table, and the dropdown list is updated dynamically.

### 2. Managing Recipients

#### Code: Handling Selection of Existing Recipients

```vbnet
Private Sub Modifiable_destinataire_AfterUpdate()
    Dim db As DAO.Database
    Dim rs As DAO.Recordset
    Dim idNom As Variant
    Dim linkValue As Variant
    Dim destinataire As String

    ' Capture the selected value
    idNom = Me.Modifiable_destinataire.Value

    ' Validate that the value is not empty
    If Trim(idNom & "") = "" Then
        MsgBox "Veuillez sélectionner une valeur ou saisir un nouveau destinataire.", vbExclamation, "Erreur"
        Exit Sub
    End If

    ' Search for the selected value in the table
    Set db = CurrentDb
    Set rs = db.OpenRecordset("SELECT * FROM [Destinataire] WHERE [N°] = " & idNom, dbOpenSnapshot)

    If Not rs.EOF Then
        ' If the value exists, fill in the labels
        Me.étiquette_destinataire_1.Caption = rs![Nom de l'entreprise]
        Me.étiquette_destinataire_2.Caption = rs![Nom du destinataire]
        Me.étiquette_destinataire_3.Caption = rs![Adresse: rue]
        Me.étiquette_destinataire_4.Caption = rs![Adresse: CP-Ville]
        Me.étiquette_destinataire_5.Caption = rs![Adresse: Pays]
        Me.étiquette_destinataire_6.Caption = rs![Numéro de téléphone]
        Me.étiquette_destinataire_7.Caption = rs![Adresse mail]

        ' Update the second dropdown (References)
        destinataire = Me.étiquette_destinataire_1.Caption
        Me.Modifiable_reference.RowSource = "SELECT Référence, Désignation FROM Reference WHERE Destinataire = '" & destinataire & "';"
        Me.Modifiable_reference.Requery

        ' Manage visibility of the special form button
        linkValue = rs![Formulaire]
        If IsNull(linkValue) Or linkValue = "" Or linkValue = "-" Then
            Me.Bouton_formulaire.Visible = False
            Me.étiquette_bouton_formulaire.Visible = False
        Else
            Me.Bouton_formulaire.Visible = True
            Me.étiquette_bouton_formulaire.Visible = True
            MsgBox "Ce fournisseur nécessite de remplir un formulaire spécial et d'envoyer un préavis par e-mail avant de soumettre la demande. Merci de compléter à la fois le formulaire Excel et le document spécial.", vbExclamation, "Erreur"
        End If
    Else
        MsgBox "Erreur : La sélection n'existe pas dans la base de données.", vbCritical, "Erreur"
        Me.Bouton_formulaire.Visible = False
        Me.étiquette_bouton_formulaire.Visible = False
    End If

    rs.Close
    Set rs = Nothing
    Set db = Nothing
End Sub
```
**Explanation:**

  * Dynamically fetches recipient details based on the selection.

  * Updates associated labels and dynamically loads references linked to the recipient.

  * Handles visibility of special form requirements based on database fields.

#### Code: Adding New Recipients

```vbnet
Private Sub Modifiable_destinataire_NotInList(NewData As String, Response As Integer)
    Dim sqlInsert As String
    Dim userResponse As Integer
    Dim nomEntreprise As String
    Dim nomDestinataire As String
    Dim adresseRue As String
    Dim adresseCPVille As String
    Dim adressePays As String
    Dim numeroTelephone As String
    Dim adresseMail As String

    ' Ask the user if they want to add the new value
    userResponse = MsgBox("Le destinataire '" & NewData & "' n'existe pas. Voulez-vous l'ajouter ?", vbYesNo + vbQuestion, "Nouveau Destinataire")

    If userResponse = vbYes Then
        ' Request additional details
        nomEntreprise = InputBox("Veuillez saisir le nom de l'entreprise :", "Nouveau Destinataire")
        nomDestinataire = InputBox("Veuillez saisir le nom du destinataire :", "Nouveau Destinataire")
        adresseRue = InputBox("Veuillez saisir l'adresse : rue :", "Nouveau Destinataire")
        adresseCPVille = InputBox("Veuillez saisir l'adresse : CP-Ville :", "Nouveau Destinataire")
        adressePays = InputBox("Veuillez saisir l'adresse : Pays :", "Nouveau Destinataire")
        numeroTelephone = InputBox("Veuillez saisir le numéro de téléphone :", "Nouveau Destinataire")
        adresseMail = InputBox("Veuillez saisir l'adresse mail :", "Nouveau Destinataire")

        ' Validate that required fields are not empty
        If Trim(nomEntreprise) = "" Or Trim(nomDestinataire) = "" Or Trim(adresseRue) = "" Then
            MsgBox "Certains champs obligatoires sont manquants. L'ajout a été annulé.", vbExclamation, "Erreur"
            Response = acDataErrContinue
            Exit Sub
        End If

        ' Insert the new record into the table
        sqlInsert = "INSERT INTO [Destinataire] ([Nom de l'entreprise], [Nom du destinataire], [Adresse: rue], [Adresse: CP-Ville], [Adresse: Pays], [Numéro de téléphone], [Adresse mail]) " & _
                    "VALUES ('" & Replace(nomEntreprise, "'", "''") & "', '" & Replace(nomDestinataire, "'", "''") & "', '" & Replace(adresseRue, "'", "''") & "', '" & Replace(adresseCPVille, "'", "''") & "', '" & Replace(adressePays, "'", "''") & "', '" & Replace(numeroTelephone, "'", "''") & "', '" & Replace(adresseMail, "'", "''") & "');"
        CurrentDb.Execute sqlInsert, dbFailOnError

        ' Update the dropdown list
        Me.Modifiable_destinataire.Undo
        Me.Modifiable_destinataire.RowSource = "SELECT [N°], [Nom de l'entreprise] FROM [Destinataire];"
        Me.Modifiable_destinataire.Requery

        ' Notify Access that the new data was successfully added
        Response = acDataErrAdded

        MsgBox "Le destinataire '" & NewData & "' a été ajouté avec succès.", vbInformation, "Succès"
    Else
        ' Cancel if the user does not want to add the value
        Response = acDataErrContinue
        MsgBox "Le destinataire n'a pas été ajouté.", vbExclamation, "Action Annulée"
    End If
End Sub
```
**Explanation:**

  * Handles the addition of new recipients dynamically.

  * Ensures all mandatory details are entered before adding to the database.

