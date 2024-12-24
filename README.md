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

### 3. Managing References

#### Code: Selecting a Reference

```vbnet
Private Sub Modifiable_reference_Click()
    Dim reference As String
    Dim id_reference As Variant

    If IsNull(Me.Modifiable_reference.Value) Then
        Me.étiquette_liste_reference.Caption = ""
    Else
        Me.étiquette_liste_reference.Caption = Me.Modifiable_reference.Value
    End If

    reference = Me.étiquette_liste_reference.Caption
    Me.Modifiable_liste_designation.RowSource = "SELECT Désignation FROM Reference WHERE Référence = '" & reference & "';"
    Me.Modifiable_liste_designation.Requery
End Sub
```
**Explanation:**

   * This subroutine triggers when a reference is selected from the dropdown list.

   * Updates the label and dynamically populates the associated designation list.

#### Code: Adding a New Reference

```vbnet
Private Sub Modifiable_liste_designation_Click()
    Dim rs As DAO.Recordset
    Dim destinataire As String
    Dim reference As String
    Dim designation As String
    Dim userResponse As Integer
    Dim sqlInsert As String

    If IsNull(Me.Modifiable_liste_designation.Value) Then
        Me.étiquette_liste_designation.Caption = ""
    Else
        Me.étiquette_liste_designation.Caption = Me.Modifiable_liste_designation.Value
    End If

    destinataire = Me.étiquette_destinataire_1.Caption
    reference = Me.étiquette_liste_reference.Caption
    designation = Me.étiquette_liste_designation.Caption

    Set rs = CurrentDb.OpenRecordset("SELECT Référence FROM Reference WHERE Référence = '" & reference & "';")

    If rs.EOF Then
        userResponse = MsgBox("La référence '" & reference & "' n'existe pas dans la base de données. Voulez-vous l'ajouter ?", vbYesNo + vbQuestion, "Nouvelle Référence")

        If userResponse = vbYes Then
            sqlInsert = "INSERT INTO Reference (Destinataire, Référence, Désignation) VALUES ('" & destinataire & "', '" & reference & "', '" & designation & "');"
            CurrentDb.Execute sqlInsert, dbFailOnError
            MsgBox "La référence a été ajoutée avec succès.", vbInformation, "Référence Ajoutée"
        Else
            MsgBox "La référence n'a pas été ajoutée.", vbExclamation, "Action Annulée"
            Me.Modifiable_reference.Value = Null
            Me.étiquette_liste_reference.Caption = ""
            rs.Close
            Set rs = Nothing
            Exit Sub
        End If
    End If
End Sub
```
**Explanation:**

   * This subroutine triggers when a user selects a designation or attempts to add a new reference.

   * Prompts the user to confirm the addition of a new reference and dynamically inserts it into the Reference table.

   * Ensures all necessary fields are correctly populated before committing to the database.
 
 ### 4. Exporting Data to Excel

 #### Code: Exporting to Excel

 ```vbnet
Private Sub Transfer_Click()
    Dim xlApp As Object
    Dim xlBook As Object
    Dim xlSheet As Object
    Dim rutaExcel As String
    Dim name As Variant

    rootExcel = "K:\A_27_TS\A_2704_TSM\A_270402_Maintenance\Formulaire d'expédition\Formulaire Standar\Modèle formulaire expéditions 2024.xlsx"

    On Error Resume Next
    Set xlApp = CreateObject("Excel.Application")
    On Error GoTo 0

    If xlApp Is Nothing Then
        MsgBox "No se pudo iniciar Excel. Asegúrate de que Excel esté instalado.", vbCritical
        Exit Sub
    End If

    Set xlBook = xlApp.Workbooks.Open(rootExcel)
    Set xlSheet = xlBook.Sheets(1)
    Dim Value(1 To 23) As Variant

    ' Collect form data to export
    Value(1) = Me.étiquette_demandeur_interne_1.Caption
    Value(2) = Me.étiquette_demandeur_interne_2.Caption
    Value(3) = Me.étiquette_demandeur_interne_3.Caption
    Value(4) = Me.étiquette_demandeur_interne_4.Caption
    Value(5) = Me.étiquette_demandeur_interne_5.Caption

    Value(6) = Me.étiquette_destinataire_1.Caption
    Value(7) = Me.étiquette_destinataire_2.Caption
    Value(8) = Me.étiquette_destinataire_3.Caption
    Value(9) = Me.étiquette_destinataire_4.Caption
    Value(10) = Me.étiquette_destinataire_5.Caption
    Value(11) = Me.étiquette_destinataire_6.Caption
    Value(12) = Me.étiquette_destinataire_7.Caption

    Value(13) = Me.étiquette_liste_reference.Caption
    Value(14) = Me.étiquette_liste_designation.Caption
    Value(15) = Me.Marchandise_3.Value
    Value(16) = Me.Marchandise_4.Value

    Value(17) = Me.livraison_1.Value
    Value(18) = Me.livraison_2.Value
    Value(19) = Me.livraison_3.Value
    Value(20) = Me.livraison_4.Value
    Value(21) = Me.livraison_5.Value

    Value(22) = Me.étiquette_date_de_livraison_1.Caption
    Value(23) = Me.date_de_livraison_1.Value

    ' Map data to Excel template
    xlSheet.Range("B7").Value = Value(1)
    xlSheet.Range("B8").Value = Value(2)
    xlSheet.Range("B9").Value = Value(3)
    xlSheet.Range("B10").Value = Value(4)
    xlSheet.Range("B11").Value = Value(5)

    xlSheet.Range("B13").Value = Value(6)
    xlSheet.Range("B14").Value = Value(7)
    xlSheet.Range("B15").Value = Value(8)
    xlSheet.Range("B16").Value = Value(9)
    xlSheet.Range("B17").Value = Value(10)
    xlSheet.Range("B18").Value = Value(11)
    xlSheet.Range("B19").Value = Value(12)

    xlSheet.Range("B21").Value = Value(13)
    xlSheet.Range("B22").Value = Value(14)
    xlSheet.Range("B23").Value = Value(15)
    xlSheet.Range("B24").Value = Value(16)

    xlSheet.Range("B26").Value = Value(17)
    xlSheet.Range("B27").Value = Value(18)
    xlSheet.Range("B28").Value = Value(19)
    xlSheet.Range("B29").Value = Value(20)
    xlSheet.Range("B30").Value = Value(21)

    xlSheet.Range("B34").Value = Value(22)
    xlSheet.Range("B35").Value = Value(23)

    ' Save the Excel file
    name = Format(Date, "yyyymmdd") & "-Formulaire d'expedition-" & Value(6) & ".xlsx"
    xlBook.SaveAs "K:\A_27_TS\A_2704_TSM\A_270402_Maintenance\Formulaire d'expédition\" & name
    xlBook.Close False
    xlApp.Quit

    MsgBox "Données exportées avec succès vers Excel.", vbInformation

    Set xlSheet = Nothing
    Set xlBook = Nothing
    Set xlApp = Nothing
End Sub
```
**Explanation:**

   * Gathers all data entered in Formulaire2 and maps it to a preformatted Excel template.

   * Saves the file with a unique name based on the date and recipient.

   * Ensures a clean workflow by releasing all Excel objects after use.

### 5. Visuals for Form 1 and Form 2
To better understand the implementation of this project, here are images of the two forms created:

#### Form 1 (Menu):
![Formulaire1](link_to_image_form1)

#### Form 2 (Detailed Form):
![Formulaire2](link_to_image_form2)

### This documentation should serve as a reference for understanding the development and functionality of the tool. Further iterations can enhance usability and extend features based on operational needs.
