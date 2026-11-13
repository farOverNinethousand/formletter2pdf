Attribute VB_Name = "formletter2pdf"
' **********************************************************************************************
' formletter2pdf
'
' Copyright 2013 Christian Bartl (www.formletter2pdf.com)
'
' Licensed under the Apache License, Version 2.0 (the "License");
' you may not use this file except in compliance with the License.
' You may obtain a copy of the License at
'
' http://www.apache.org/licenses/LICENSE-2.0
'
' Unless required by applicable law or agreed to in writing, software
' distributed under the License is distributed on an "AS IS" BASIS,
' WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
' See the License for the specific language governing permissions and
' limitations under the License.
'
' This notice must appear in all copies of the script.
' **********************************************************************************************

Option Explicit

Public Const AppVersion As String = "2.11"

Public DefaultFiletype As String
Public DefaultFilelocation As String
Public DefaultFoldername As String
Public DefaultFilenamepattern As String
Public CustomDefaultFilenamePattern As Variant
Public DefaultExistingFiles As String
Public DefaultWriteLog As Boolean
Public DefaultAbortIfEmptyField As String
Public DefaultStopAtEmptyRecord As Boolean
Public DefaultOpenFolder As String
Public DefaultConfirmBeforeWriting As String
Public DefaultResultMessage As String
Public DefaultQuickScan As Boolean
Public DefaultLanguage As String
Public DefaultSecondsPerPage As Double

Global StorageFiletype As String
Global StorageFilelocation As String
Global StorageFoldername As String
Global StorageFilenamepattern As String
Global StorageExistingFiles As String
Global StorageWriteLog As String
Global StorageAbortIfEmpty As String
Global StorageAbortIfEmptyField As String
Global StorageStopAtEmptyRecord As String
Global StorageOpenFolder As String
Global StorageConfirmBeforeWriting As String
Global StorageResultMessage As String
Global StorageQuickScan As String
Global StorageFormWidth As Single
Global StorageFormHeight As Single
Global StorageLanguage As String

Public Const LangEnglish As String = "en"
Public Const LangGerman As String = "de"

' language which is used for all texts: "en" or "de"
Public CurrentLanguage As String

Private Texts As Collection


Sub formletter2pdf()

    '
    ' define default values
    '

        ' required; one or more of "pdf", "docx", "PDF/A (ISO 19005-1-compatible, .pdf)", "Word 97-2003 Document (.doc)", "Single File Web Page (.mht)", "Web Page (.htm)", "Web Page, Filtered (.htm)", "Rich Text Format (.rtf)", "Plain Text (.txt)"
        ' notes: separate multiple file types with ";" (e.g. "pdf;docx"). "PDF/A (ISO 19005-1-compatible, .pdf)" is only available on Office for Windows
    DefaultFiletype = "pdf"

        ' optional; File location can be 'Desktop' or a full path (e.g. D:\Foldername) without trailing slash, where the folder with the generated files will be stored.
    DefaultFilelocation = "Desktop"

        ' optional; Use the {%timestamp%} marker in the folder name to prevent overwriting if formletter2pdf is executed multiple times.
    DefaultFoldername = "formletter-{%timestamp%}"

        ' optional; Make sure to create unique file names by including the values of your data source fields like {field_name} or a generic counter {%counter%}. The file name of the word document can be included via {%filename%}.
    DefaultFilenamepattern = "{%counter%}-formletter-{last_name}-{ID}"
        ' note: the pattern "{%counter%}-formletter-{last_name}-{ID}" which is shipped with formletter2pdf uses fields which normally do not exist in your data source. If it is still used in the dialog, formletter2pdf shows an error and does not start. To use your own default pattern without getting this error, set CustomDefaultFilenamePattern below.

        ' optional; Your own default file name pattern, e.g. "{%counter%}-letter-{Name}". If this is set (not Null and not empty), it is shown in the dialog instead of DefaultFilenamepattern. Because it is not the pattern shipped with formletter2pdf, it does not trigger the error described above. Leave it empty ("" or Null) to use DefaultFilenamepattern.
    CustomDefaultFilenamePattern = ""

        ' optional; What to do if a file with the same name already exists: "ask", "overwrite", "skip" or "number" (auto rename: saves the file as "name (2).pdf", "name (3).pdf", ...)
    DefaultExistingFiles = "ask"

        ' optional; True writes a log file (formletter2pdf-log.txt) into the folder with the generated files.
    DefaultWriteLog = False

        ' optional; Name of a data source field without braces (e.g. "last_name"). If set, the export stops at the first record where this field is empty; all form letters before are saved.
    DefaultAbortIfEmptyField = ""

        ' optional; True (default) stops reading the data source and the export at the first record where ALL fields are empty (e.g. empty rows at the end of an Excel sheet). If False, empty records are skipped.
    DefaultStopAtEmptyRecord = True

        ' optional; Open the folder with the generated files in Windows Explorer after the export (Windows only): "ask", "open" or "no"
    DefaultOpenFolder = "ask"

        ' optional; Confirm dialog with the number of files and the target folder before any file is written: "always", "warnings" (default; only if there is something to warn about, e.g. the export stops early or file names already exist) or "never"
    DefaultConfirmBeforeWriting = "warnings"

        ' optional; Message after the export with the result: "always" (default), "warnings" (only if the export was cancelled or stopped early, or files were skipped) or "never". Errors are always shown.
    DefaultResultMessage = "always"

        ' optional; True (default) limits the number of records read by the scan button to 2 (quick check of the file names). The export always reads all records.
    DefaultQuickScan = True

        ' optional; Language of the dialog: "auto" (default, the language of Microsoft Word is used: german if Word is german, otherwise english), "en" or "de". The language can also be changed in the dialog.
    DefaultLanguage = "auto"

        ' optional; Estimated time in seconds to create one page of a form letter (all file types together). The number of pages of the main document multiplied with this value is the estimate per form letter, used for the remaining time shown during the export until real times are measured. The estimate is adjusted with every created form letter.
    DefaultSecondsPerPage = 0.5



    '
    ' test if document is a formletter and view form
    '

    Dim MailMergeState As Object

    ' language of the dialog and the messages
    InitLanguage StorageLanguage, DefaultLanguage

    Set MailMergeState = ActiveDocument.MailMerge
    If MailMergeState.State = wdMainAndDataSource Then
        formletter2pdf_form.Show
    Else
        MsgBox Tx("X_NO_FORMLETTER"), vbOKOnly + vbCritical
    End If

End Sub


' ***********************************************************************************************
' Texts and language handling
' ***********************************************************************************************

' All texts of the dialog in all languages. A text is looked up with Tx("KEY"); %1, %2, ... in a text are
' replaced with the arguments of TxF("KEY", arg1, arg2, ...). TxC(Number, "KEY", ...) selects the text
' KEY_1 (Number = 1) or KEY_N and replaces %1 with the Number.
'
' To add a language: add a column to every Add line, extend LanguageIndex and DetectLanguage.
' The source files must stay plain ASCII, so special characters are written as escape sequences:
'   \u00e4 = a-umlaut (any \uXXXX with the hex code of the character), \n = new line, \q = double quote


Public Function DetectLanguage() As String
    ' language of the user interface of word; german for all german variants (de-DE, de-AT, de-CH, ...), otherwise english
    Dim LanguageId As Long

    DetectLanguage = LangEnglish
    On Error Resume Next
    LanguageId = Application.LanguageSettings.LanguageID(msoLanguageIDUI)
    If Err.Number = 0 Then
        ' the lower 10 bits are the primary language, 7 = german
        If (LanguageId And &H3FF) = 7 Then DetectLanguage = LangGerman
    End If
    On Error GoTo 0
End Function


Public Sub InitLanguage(ByVal StoredLanguage As String, ByVal DefaultLang As String)
    ' language at the start: the language chosen in the dialog before, otherwise the default of the module
    ' formletter2pdf ("en" or "de"), otherwise (default "auto") the language of word
    If StoredLanguage = LangEnglish Or StoredLanguage = LangGerman Then
        SetLanguage StoredLanguage
    ElseIf LCase(DefaultLang) = LangEnglish Or LCase(DefaultLang) = LangGerman Then
        SetLanguage LCase(DefaultLang)
    Else
        SetLanguage DetectLanguage()
    End If
End Sub


Public Sub SetLanguage(ByVal Language As String)
    If Language = LangGerman Then
        CurrentLanguage = LangGerman
    Else
        CurrentLanguage = LangEnglish
    End If
    If Texts Is Nothing Then LoadTexts
End Sub


Private Function LanguageIndex() As Long
    ' position of the text of the current language in an Add line (0 = english)
    If CurrentLanguage = LangGerman Then
        LanguageIndex = 1
    Else
        LanguageIndex = 0
    End If
End Function


Public Function Tx(ByVal Key As String) As String
    ' text of the key in the current language; english if the text is not translated; [KEY] if the key is unknown
    Dim Entry As Variant

    If Texts Is Nothing Then LoadTexts
    On Error Resume Next
    Entry = Texts(Key)
    If Err.Number <> 0 Then
        Tx = "[" & Key & "]"
        On Error GoTo 0
        Exit Function
    End If
    On Error GoTo 0

    Tx = Entry(LanguageIndex())
    If Tx = "" Then Tx = Entry(0)
End Function


Public Function TxF(ByVal Key As String, ParamArray Args() As Variant) As String
    ' text with the arguments inserted for %1, %2, ...
    Dim Values() As Variant
    Dim i As Long

    ReDim Values(0 To UBound(Args) + 1)
    For i = 0 To UBound(Args)
        Values(i) = Args(i)
    Next i
    TxF = Subst(Tx(Key), Values, UBound(Args))
End Function


Public Function TxC(ByVal Number As Long, ByVal Key As String, ParamArray Args() As Variant) As String
    ' text for a number: KEY_1 for one, KEY_N otherwise; %1 is the number, %2, %3, ... are the further arguments
    Dim Values() As Variant
    Dim i As Long

    ReDim Values(0 To UBound(Args) + 1)
    Values(0) = Number
    For i = 0 To UBound(Args)
        Values(i + 1) = Args(i)
    Next i
    TxC = Subst(Tx(Key & IIf(Number = 1, "_1", "_N")), Values, UBound(Args) + 1)
End Function


Private Function Subst(ByVal Text As String, Values() As Variant, ByVal LastIndex As Long) As String
    ' replaces %1 to %9 in one pass, so a value which contains "%2" is not replaced again
    Dim Result As String
    Dim Digit As Long
    Dim i As Long

    i = 1
    Do While i <= Len(Text)
        Digit = 0
        If Mid(Text, i, 1) = "%" And i < Len(Text) Then
            If Mid(Text, i + 1, 1) Like "[1-9]" Then Digit = CLng(Mid(Text, i + 1, 1))
        End If

        If Digit > 0 And Digit - 1 <= LastIndex Then
            Result = Result & CStr(Values(Digit - 1))
            i = i + 2
        Else
            Result = Result & Mid(Text, i, 1)
            i = i + 1
        End If
    Loop
    Subst = Result
End Function


Private Function Decode(ByVal Text As String) As String
    Dim p As Long

    Text = Replace(Text, "\n", vbNewLine)
    Text = Replace(Text, "\q", Chr(34))
    p = InStr(Text, "\u")
    Do While p > 0
        Text = Left(Text, p - 1) & ChrW(CLng("&H" & Mid(Text, p + 2, 4))) & Mid(Text, p + 6)
        p = InStr(Text, "\u")
    Loop
    Decode = Text
End Function


Private Sub Add(ByVal Key As String, ByVal English As String, ByVal German As String)
    Texts.Add Array(Decode(English), Decode(German)), Key
End Sub


' ***********************************************************************************************
' Text table (english, german); keep at the end of the module
' ***********************************************************************************************

Private Sub LoadTexts()
    Set Texts = New Collection

    ' ---------------------------------------------------------------------------------------
    ' original controls of the form layout, their texts are replaced when the language is set
    ' ---------------------------------------------------------------------------------------
    Add "L_FILETYPE", "Export files as", "Exportieren als"
    Add "L_SUBHEAD", "Updates & documentation on github.com/farOverNinethousand/formletter2pdf", "Updates & Dokumentation auf github.com/farOverNinethousand/formletter2pdf"
    Add "L_OLDSITE", "Old website: formletter2pdf.com", "Alte Webseite: formletter2pdf.com"
    Add "L_FOLDER", "Folder name pattern", "Ordnername-Muster"
    Add "L_FOLDER_HINT", "Use the {%timestamp%} marker in the folder name to prevent overwriting if formletter2pdf is executed multiple times.", "Mit dem Platzhalter {%timestamp%} im Ordnernamen wird nichts \u00fcberschrieben, wenn formletter2pdf mehrfach ausgef\u00fchrt wird."
    Add "L_FILE", "File name pattern", "Dateiname-Muster"
    Add "L_FILE_HINT", "Make sure to create unique file names by including the values of your data source fields like {field_name} or a generic counter {%counter%}. The file name of the word document can be included via {%filename%}.", "Sorge f\u00fcr eindeutige Dateinamen: Nutze Werte deiner Datenquellenfelder wie {field_name} oder einen Z\u00e4hler {%counter%}. Der Dateiname des Word-Dokuments l\u00e4sst sich mit {%filename%} einf\u00fcgen."
    Add "L_FILE_HINT_ACTION", "Click to list all available data source field markers.", "Alle verf\u00fcgbaren Platzhalter anzeigen."
    Add "L_SAVETO", "Save to", "Speichern in"
    Add "L_SAVETO_HELP", "Choose the location where the folder with the generated files should be stored.", "W\u00e4hle den Ort, an dem der Ordner mit den erzeugten Dateien gespeichert werden soll."
    Add "B_CHOOSE", "Choose directory ...", "Ordner w\u00e4hlen ..."
    Add "L_PRINT", "And also print them directly", "Und zus\u00e4tzlich direkt drucken"
    Add "L_PRINT_HINT", "The print dialog to choose printer and print settings is shown once at the beginning of the export.", "Der Druckdialog zur Auswahl von Drucker und Druckeinstellungen erscheint einmal zu Beginn des Exports."
    Add "L_DONATE", "Support development with your donation!", "Unterst\u00fctze die Entwicklung mit einer Spende!"

    ' ---------------------------------------------------------------------------------------
    ' buttons, frames and labels
    ' ---------------------------------------------------------------------------------------
    Add "L_LANGUAGE", "Language:", "Sprache:"
    Add "B_CLOSE", "Close", "Schlie\u00dfen"
    Add "B_SAVE_DEFAULT", "Export files", "Export starten"
    Add "B_SAVE_N", "Export %1", "%1 exportieren"
    Add "B_ABORT_SAVE", "Abort export", "Exportieren abbrechen"
    Add "B_SCAN", "Scan", "Scannen"
    Add "B_REFRESH", "Refresh", "Aktualisieren"
    Add "B_ABORT_SCAN", "Abort scan", "Scan abbrechen"
    Add "B_ABORT", "Abort", "Abbrechen"
    Add "F_OPTIONS", "Additional options", "Weitere Optionen"
    Add "F_PREVIEW", "Preview", "Vorschau"
    Add "F_CONFLICT", "File already exists", "Datei existiert bereits"
    Add "F_PROGRESS", "Progress", "Fortschritt"
    Add "F_MESSAGES", "Messages", "Meldungen"

    ' ---------------------------------------------------------------------------------------
    ' options and settings
    ' ---------------------------------------------------------------------------------------
    Add "L_EXISTING", "If a file already exists:", "Wenn eine Datei bereits existiert:"
    Add "CBO_EX_ASK", "Ask me", "Nachfragen"
    Add "CBO_EX_OVERWRITE", "Overwrite", "\u00dcberschreiben"
    Add "CBO_EX_SKIP", "Skip", "\u00dcberspringen"
    Add "CBO_EX_NUMBER", "Auto rename", "Automatisch umbenennen"
    Add "MODE_ASK", "ask", "nachfragen"
    Add "MODE_OVERWRITE", "overwrite", "\u00fcberschreiben"
    Add "MODE_SKIP", "skip", "\u00fcberspringen"
    Add "MODE_NUMBER", "auto rename", "automatisch umbenennen"
    Add "O_LOG", "Write a log file (%1) into the target folder", "Protokoll (%1) in den Zielordner schreiben"
    Add "O_STOP_ALL", "Stop export/scan at first record with ALL fields empty (faster with many empty rows)", "Export/Scan beim ersten Datensatz stoppen, bei dem ALLE Felder leer sind (schneller bei vielen leeren Zeilen)"
    Add "O_STOP_FIELD", "Stop export/scan at first record with following empty field:", "Export/Scan beim ersten Datensatz stoppen, wenn Feld leer:"
    Add "O_QUICK", "Enable quick scan (limits number of records to scan to %1)", "Schnellscan aktivieren (begrenzt die gescannten Datens\u00e4tze auf %1)"
    Add "O_APPLY_ALL", "Apply to all further existing files", "Auf alle weiteren vorhandenen Dateien anwenden"
    Add "L_MSG_CONFIRM", "Confirm before writing files:", "Vor dem Schreiben best\u00e4tigen:"
    Add "L_MSG_RESULT", "Result message after export:", "Ergebnismeldung nach dem Export:"
    Add "L_MSG_OPEN", "Open output folder on success:", "Ausgabeordner nach Erfolg \u00f6ffnen:"
    Add "M_ALWAYS", "Always", "Immer"
    Add "M_WARNINGS", "Only when warnings exist", "Nur bei Warnungen"
    Add "M_NEVER", "Never", "Nie"
    Add "M_ASK", "Ask", "Nachfragen"
    Add "M_OPEN", "Open folder", "Ordner \u00f6ffnen"
    Add "M_NOOPEN", "Do not open folder", "Ordner nicht \u00f6ffnen"

    ' ---------------------------------------------------------------------------------------
    ' numbers with a noun (%1 is the number)
    ' ---------------------------------------------------------------------------------------
    Add "N_FILE_1", "%1 file", "%1 Datei"
    Add "N_FILE_N", "%1 files", "%1 Dateien"
    Add "N_LETTER_1", "%1 form letter", "%1 Serienbrief"
    Add "N_LETTER_N", "%1 form letters", "%1 Serienbriefe"
    Add "N_EMPTYREC_1", "%1 empty record", "%1 leerer Datensatz"
    Add "N_EMPTYREC_N", "%1 empty records", "%1 leere Datens\u00e4tze"
    Add "N_EXISTFILE_1", "%1 existing file", "%1 vorhandene Datei"
    Add "N_EXISTFILE_N", "%1 existing files", "%1 vorhandene Dateien"

    ' ---------------------------------------------------------------------------------------
    ' preview, status and progress
    ' ---------------------------------------------------------------------------------------
    Add "S_SCAN_HINT", "Click \q%1\q to read the data source and show the file names.", "Klicke auf \q%1\q, um die Datenquelle zu lesen und die Dateinamen anzuzeigen."
    Add "S_SCAN_AGAIN", "Click \q%1\q to read the records of the data source again.", "Klicke auf \q%1\q, um die Datens\u00e4tze der Datenquelle erneut zu lesen."
    Add "S_SCAN_ABORTED", "Scan aborted. Click \q%1\q to read the data source.", "Scan abgebrochen. Klicke auf \q%1\q, um die Datenquelle zu lesen."
    Add "S_READING", "Reading the data source, please wait ...", "Datenquelle wird gelesen, bitte warten ..."
    Add "PR_READING", "Reading the data source ...", "Datenquelle wird gelesen ..."
    Add "PR_READING_REC", "Reading the data source: record %1 of %2", "Datenquelle wird gelesen: Datensatz %1 von %2"
    Add "PR_NOT_STARTED", "Not started", "Nicht gestartet"
    Add "PR_STARTING", "Starting ...", "Starte ..."
    Add "PR_LETTER", "Form letter %1 of %2: %3", "Serienbrief %1 von %2: %3"
    Add "PR_CANCEL_SCAN", "Cancelling the scan ...", "Scan wird abgebrochen ..."
    Add "PR_CANCEL_SAVE", "Cancelling after the current form letter ...", "Abbruch nach dem aktuellen Serienbrief ..."
    Add "ETA_REMAINING", "about %1 remaining", "noch ca. %1"
    Add "DUR_S", "%1 s", "%1 s"
    Add "DUR_MS", "%1 min %2 s", "%1 Min. %2 s"
    Add "DUR_HM", "%1 h %2 min", "%1 Std. %2 Min."
    Add "P_SELECT_ROW", "Select a row to see the complete file name.", "Zeile ausw\u00e4hlen, um den vollst\u00e4ndigen Dateinamen zu sehen."
    Add "P_RECORD", "Record %1", "Datensatz %1"
    Add "REC_REF", "record %1/%2", "Datensatz %1/%2"
    Add "REC_REF_SHORT", "record %1", "Datensatz %1"
    Add "P_ALL_EMPTY_STOP", "All fields are empty: export stops here at %1", "Alle Felder sind leer: Export endet hier bei %1"
    Add "P_EMPTY_SKIPPED", "Empty record, skipped", "Leerer Datensatz, \u00fcbersprungen"
    Add "P_FIELD_STOP", "{%1} is empty: export stops here at %2", "{%1} ist leer: Export endet hier bei %2"
    Add "P_NOT_EXPORTED", "Not exported (export stopped before)", "Nicht exportiert (Export vorher beendet)"
    Add "P_RESERVED", "Reserved Windows name", "Reservierter Windows-Name"
    Add "P_PATH_LONG", "Path too long", "Pfad zu lang"
    Add "P_DUP", "Duplicate name: %1", "Doppelter Name: %1"
    Add "P_EXISTS", "File exists: %1", "Datei existiert: %1"
    Add "P_QUICK_ROW", "Quick scan is active: only the first %1 of %2 records are shown. Switch off quick scan and scan again to see all of them.", "Schnellscan ist aktiv: Nur die ersten %1 von %2 Datens\u00e4tzen werden angezeigt. Schnellscan ausschalten und neu scannen, um alle zu sehen."
    Add "P_SUM_CREATED", "%1, %2 will be created.", "%1, %2 werden erstellt."
    Add "P_QUICK_SUM", "Quick scan: first %1 of %2 records only. ", "Schnellscan: nur die ersten %1 von %2 Datens\u00e4tzen. "
    Add "P_STOPS_FIELD", " Stops at %1 because {%2} is empty.", " Endet bei %1, weil {%2} leer ist."
    Add "P_STOPS_FIELD_FOUND", " Stops at %1 because {%2} is empty (%3 found).", " Endet bei %1, weil {%2} leer ist (%3 gefunden)."
    Add "P_STOPS_EMPTY", " Stops at %1 (all fields empty).", " Endet bei %1 (alle Felder leer)."
    Add "P_CONFLICT_1", " %1 name already exists (%2).", " %1 Name existiert bereits (%2)."
    Add "P_CONFLICT_N", " %1 names already exist (%2).", " %1 Namen existieren bereits (%2)."
    Add "P_EMPTYSKIP_1", " %1 empty record is skipped.", " %1 leerer Datensatz wird \u00fcbersprungen."
    Add "P_EMPTYSKIP_N", " %1 empty records are skipped.", " %1 leere Datens\u00e4tze werden \u00fcbersprungen."
    Add "P_UNKNOWN_MARKER", "Unknown marker in file name pattern: %1", "Unbekannter Platzhalter im Dateiname-Muster: %1"

    ' ---------------------------------------------------------------------------------------
    ' checks of the settings (errors)
    ' ---------------------------------------------------------------------------------------
    Add "E_SHIPPED", "The file name pattern is still the pattern which is shipped with formletter2pdf:\n%1\n\nIt uses fields which normally do not exist in your data source. Please enter your own file name pattern with the fields of your data source (the available fields are listed with the link below the file name pattern).", "Das Dateiname-Muster ist noch das mit formletter2pdf mitgelieferte Muster:\n%1\n\nEs verwendet Felder, die in deiner Datenquelle normalerweise nicht existieren. Bitte gib ein eigenes Dateiname-Muster mit den Feldern deiner Datenquelle ein (die verf\u00fcgbaren Felder zeigt der Link unter dem Dateiname-Muster)."
    Add "E_NO_FILETYPE", "Please choose at least one file type.", "Bitte w\u00e4hle mindestens einen Dateityp."
    Add "E_DUP_TYPES", "You have selected several file types which produce the same file extension, the files would overwrite each other.\n\nPlease select only one file type per file extension. Selected more than once:\n\n%1", "Du hast mehrere Dateitypen gew\u00e4hlt, die dieselbe Dateiendung erzeugen, die Dateien w\u00fcrden sich gegenseitig \u00fcberschreiben.\n\nBitte w\u00e4hle pro Dateiendung nur einen Dateityp. Mehrfach gew\u00e4hlt:\n\n%1"
    Add "E_NO_PATTERN", "Please provide a file name pattern.", "Bitte gib ein Dateiname-Muster an."
    Add "E_NO_FIELD", "Please choose the data source field which must not be empty.", "Bitte w\u00e4hle das Datenquellenfeld, das nicht leer sein darf."
    Add "E_NO_LOCATION", "Please provide a storage location in \q%1\q where the generated files can be saved to.", "Bitte gib unter \q%1\q einen Speicherort an, in dem die erzeugten Dateien gespeichert werden k\u00f6nnen."
    Add "E_LOC_MISSING", "The storage location does not exist:\n%1\n\nPlease choose an existing folder in \q%2\q.", "Der Speicherort existiert nicht:\n%1\n\nBitte w\u00e4hle unter \q%2\q einen vorhandenen Ordner."
    Add "E_LOC_WRITE", "You do not have write access to the storage location:\n%1\n\nPlease choose another storage location in \q%2\q.", "Du hast keinen Schreibzugriff auf den Speicherort:\n%1\n\nBitte w\u00e4hle unter \q%2\q einen anderen Speicherort."
    Add "E_NO_FOLDER", "Please provide a folder name pattern.", "Bitte gib ein Ordnername-Muster an."
    Add "E_FOLDER_CHARS", "The folder name pattern must not contain any of the following characters: %1", "Das Ordnername-Muster darf keines der folgenden Zeichen enthalten: %1"
    Add "E_RESERVED_1", "%1 file name is reserved by Windows and cannot be used:", "%1 Dateiname ist von Windows reserviert und kann nicht verwendet werden:"
    Add "E_RESERVED_N", "%1 file names are reserved by Windows and cannot be used:", "%1 Dateinamen sind von Windows reserviert und k\u00f6nnen nicht verwendet werden:"
    Add "E_RESERVED_HINT", "Please change the file name pattern, e.g. by adding a prefix or the counter {%counter%}.", "Bitte \u00e4ndere das Dateiname-Muster, z. B. mit einem Pr\u00e4fix oder dem Z\u00e4hler {%counter%}."
    Add "E_PATH_1", "%1 file path would be too long for Windows (maximum %2 characters).", "%1 Dateipfad w\u00e4re f\u00fcr Windows zu lang (maximal %2 Zeichen)."
    Add "E_PATH_N", "%1 file paths would be too long for Windows (maximum %2 characters).", "%1 Dateipfade w\u00e4ren f\u00fcr Windows zu lang (maximal %2 Zeichen)."
    Add "E_PATH_INFO", "The longest path has %1 characters (record %2): %3\n\nPlease choose a shorter storage location or folder name or a shorter file name pattern.", "Der l\u00e4ngste Pfad hat %1 Zeichen (Datensatz %2): %3\n\nBitte w\u00e4hle einen k\u00fcrzeren Speicherort oder Ordnernamen oder ein k\u00fcrzeres Dateiname-Muster."
    Add "E_NOTHING_FIELD", "Nothing to export: the field {%1} is already empty in the first record (record %2).", "Nichts zu exportieren: Das Feld {%1} ist schon im ersten Datensatz leer (Datensatz %2)."
    Add "E_NO_RECORDS", "There is no record with data to export.", "Es gibt keinen Datensatz mit Daten zum Exportieren."

    ' ---------------------------------------------------------------------------------------
    ' questions before and after the export
    ' ---------------------------------------------------------------------------------------
    Add "Q_FOLDER_EXISTS", "A folder with the given name already exists at the storage location:\n%1\n\nDo you want to export the files into this folder? Files which already exist are handled according to the option \q%2 %3\q.", "Am Speicherort existiert bereits ein Ordner mit diesem Namen:\n%1\n\nSollen die Dateien in diesen Ordner exportiert werden? Bereits vorhandene Dateien werden gem\u00e4\u00df der Option \q%2 %3\q behandelt."
    Add "C_ATTN_A", "ATTENTION: The field {%1} is empty in %2. The export stops there, the following records are not exported.", "ACHTUNG: Das Feld {%1} ist in %2 leer. Der Export endet dort, die folgenden Datens\u00e4tze werden nicht exportiert."
    Add "C_ATTN_B", "ATTENTION: The field {%1} is empty in %2. The export stops there, so only %3 of %4 will be exported.", "ACHTUNG: Das Feld {%1} ist in %2 leer. Der Export endet dort, deshalb werden nur %3 von %4 exportiert."
    Add "C_CONFLICTS_1", "%1 file name already exists or is used twice (%2).", "%1 Dateiname existiert bereits oder wird doppelt verwendet (%2)."
    Add "C_CONFLICTS_N", "%1 file names already exist or are used twice (%2).", "%1 Dateinamen existieren bereits oder werden doppelt verwendet (%2)."
    Add "C_EMPTYSTOP", "The export stops at %1 because all fields are empty there. Further records are not exported.", "Der Export endet bei %1, weil dort alle Felder leer sind. Weitere Datens\u00e4tze werden nicht exportiert."
    Add "C_SUMMARY", "%1 (%2) will be exported to:\n%3\n\nThe export can take a couple of minutes, depending on the number of files.", "%1 (%2) werden exportiert nach:\n%3\n\nDer Export kann je nach Anzahl der Dateien einige Minuten dauern."
    Add "C_MAC_NEW", "Microsoft Word will be UNUSABLE during this time and MAY FEEL LIKE the programm is FROZEN.", "Microsoft Word ist in dieser Zeit NICHT BENUTZBAR und WIRKT WIE EINGEFROREN."
    Add "C_MAC_OLD", "The formletter2pdf window will be hidden during this time and Microsoft Word will open and close new document windows visually.", "Das formletter2pdf-Fenster wird in dieser Zeit ausgeblendet, und Microsoft Word \u00f6ffnet und schlie\u00dft sichtbar neue Dokumentfenster."
    Add "C_WIN", "The progress is shown in this dialog. You can stop the process with the cancel button.", "Der Fortschritt wird in diesem Dialog angezeigt. Mit der Schaltfl\u00e4che zum Abbrechen kannst du den Vorgang stoppen."
    Add "C_CONTINUE", "Want to continue?", "Fortfahren?"
    Add "CONFLICT_TEXT", "Already existing in the target folder:\n%1\n\nWhat should formletter2pdf do?", "Im Zielordner bereits vorhanden:\n%1\n\nWas soll formletter2pdf tun?"
    Add "CONFLICT_MAC", "Already existing: %1\n\nYes = overwrite\nNo = skip\nCancel = auto rename", "Bereits vorhanden: %1\n\nJa = \u00fcberschreiben\nNein = \u00fcberspringen\nAbbrechen = automatisch umbenennen"
    Add "Q_OPEN", "Do you like to open the target saving directory in Windows Explorer now?", "M\u00f6chtest du den Zielordner jetzt im Windows-Explorer \u00f6ffnen?"

    ' ---------------------------------------------------------------------------------------
    ' result of the export
    ' ---------------------------------------------------------------------------------------
    Add "R_LETTERS", "%1 exported (%2).", "%1 exportiert (%2)."
    Add "R_OVERWRITTEN", "%1 overwritten.", "%1 \u00fcberschrieben."
    Add "R_NUMBERED", "%1 saved with a new name (auto rename) because the name already existed.", "%1 unter neuem Namen gespeichert (automatisch umbenannt), weil der Name schon existierte."
    Add "R_SKIPPED", "%1 skipped because the name already existed.", "%1 \u00fcbersprungen, weil der Name schon existierte."
    Add "R_EMPTY", "%1 skipped.", "%1 \u00fcbersprungen."
    Add "R_DURATION", "Duration: %1", "Dauer: %1"
    Add "R_LOG", "Log file: %1", "Protokolldatei: %1"
    Add "M_CANCELLED", "The export has been cancelled.", "Der Export wurde abgebrochen."
    Add "M_STOPPED_EMPTY", "The export stopped at %1 because all fields are empty there. Further records were not exported.", "Der Export endete bei %1, weil dort alle Felder leer sind. Weitere Datens\u00e4tze wurden nicht exportiert."
    Add "M_STOPPED_FIELD", "The export stopped at %1 because the field {%2} is empty.", "Der Export endete bei %1, weil das Feld {%2} leer ist."
    Add "M_STOPPED_FIELD_A", "The following records were not exported.", "Die folgenden Datens\u00e4tze wurden nicht exportiert."
    Add "M_STOPPED_FIELD_B", "Only %1 of %2 have been exported.", "Nur %1 von %2 wurden exportiert."
    Add "M_OK", "Export finished successfully!", "Export erfolgreich abgeschlossen!"

    ' ---------------------------------------------------------------------------------------
    ' errors during the export and while reading the data source
    ' ---------------------------------------------------------------------------------------
    Add "X_FOLDER_RIGHTS", "You do not have proper file system rights on the given storage location. Please choose another storage location.", "Du hast keine ausreichenden Dateisystemrechte auf dem angegebenen Speicherort. Bitte w\u00e4hle einen anderen Speicherort."
    Add "X_FOLDER_INVALID", "The given location isn't a valid storage location. Please choose another storage location.", "Der angegebene Ort ist kein g\u00fcltiger Speicherort. Bitte w\u00e4hle einen anderen Speicherort."
    Add "X_FOLDER_CREATE", "The folder for the generated files could not be created: Error %1 - %2", "Der Ordner f\u00fcr die erzeugten Dateien konnte nicht angelegt werden: Fehler %1 - %2"
    Add "X_READ", "The data source could not be read: Error %1 - %2", "Die Datenquelle konnte nicht gelesen werden: Fehler %1 - %2"
    Add "X_NO_DOC", "The mail merge did not create a document.", "Der Seriendruck hat kein Dokument erzeugt."
    Add "X_CONTEXT", "Record %1, file: %2", "Datensatz %1, Datei: %2"
    Add "X_4198", "formletter2pdf has been aborted because an invalid file name has been detected. Please check the values of the fields used in the file name pattern for invalid characters and run formletter2pdf again.", "formletter2pdf wurde abgebrochen, weil ein ung\u00fcltiger Dateiname erkannt wurde. Bitte pr\u00fcfe die Werte der im Dateiname-Muster verwendeten Felder auf ung\u00fcltige Zeichen und starte formletter2pdf erneut."
    Add "X_5852", "formletter2pdf could not export the files because this document isn't a form letter. Please start the mail merge function first and then run formletter2pdf again.", "formletter2pdf konnte die Dateien nicht exportieren, weil dieses Dokument kein Serienbrief ist. Bitte starte zuerst den Seriendruck und f\u00fchre formletter2pdf dann erneut aus."
    Add "X_70", "The file could not be saved. Please check that you have write access to the storage location and that the file is not opened in another program.", "Die Datei konnte nicht gespeichert werden. Bitte pr\u00fcfe, ob du Schreibzugriff auf den Speicherort hast und die Datei nicht in einem anderen Programm ge\u00f6ffnet ist."
    Add "X_76", "The given location isn't a valid storage location. Please choose another storage location.", "Der angegebene Ort ist kein g\u00fcltiger Speicherort. Bitte w\u00e4hle einen anderen Speicherort."
    Add "X_UNKNOWN", "Unknown Error: %1 - %2%3\n\nPlease try to run formletter2pdf again.", "Unbekannter Fehler: %1 - %2%3\n\nBitte versuche, formletter2pdf erneut auszuf\u00fchren."

    ' ---------------------------------------------------------------------------------------
    ' start of the macro
    ' ---------------------------------------------------------------------------------------
    Add "X_NO_FORMLETTER", "formletter2pdf cannot be executed because this document isn't a form letter. Please start the mail merge function first and then run formletter2pdf again.", "formletter2pdf kann nicht ausgef\u00fchrt werden, weil dieses Dokument kein Serienbrief ist. Bitte starte zuerst den Seriendruck und f\u00fchre formletter2pdf dann erneut aus."
End Sub
