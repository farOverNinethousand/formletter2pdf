VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} formletter2pdf_form 
   Caption         =   "formletter2pdf"
   ClientHeight    =   9000.001
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   7605
   OleObjectBlob   =   "formletter2pdf_form.frx":0000
   StartUpPosition =   2  'Bildschirmmitte
End
Attribute VB_Name = "formletter2pdf_form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
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

#If Mac Then
    ' no sleep function and no window functions needed on mac
#ElseIf VBA7 Then
    Private Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    Private Declare PtrSafe Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As LongPtr
    Private Declare PtrSafe Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As LongPtr, ByVal nIndex As Long) As Long
    Private Declare PtrSafe Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As LongPtr, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare PtrSafe Function DrawMenuBar Lib "user32" (ByVal hWnd As LongPtr) As Long
#Else
    Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
    Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
    Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
    Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
    Private Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long
#End If

Private Const TimestampFormat As String = "dd.mm.yyyy-hh.mm.ss"
Private Const LogFileName As String = "formletter2pdf-log.txt"
Private Const MaxFileNameLength As Long = 150
Private Const MaxPathLength As Long = 259
Private Const PathNumberReserve As Long = 6
Private Const ExistingModes As String = "ask,overwrite,skip,number"
Private Const EtaPriorWeight As Long = 2
Private Const PanelTop As Single = 32
Private Const OpenFolderModes As String = "ask,open,no"
Private Const MessageModes As String = "always,warnings,never"
Private Const MessagesFrameHeight As Single = 92
Private Const QuickScanLimit As Long = 2
Private Const ShippedFilenamepattern As String = "{%counter%}-formletter-{last_name}-{ID}"
Private Const PanelWidth As Single = 300
Private Const PreviewMinColumnWidths As String = "32 pt;260 pt;200 pt"
Private Const FiletypeRowHeight As Single = 17

' additional controls, created at runtime (the layout in the .frx stays untouched)
Private WithEvents chkFt1 As MSForms.CheckBox
Private WithEvents chkFt2 As MSForms.CheckBox
Private WithEvents chkFt3 As MSForms.CheckBox
Private WithEvents chkFt4 As MSForms.CheckBox
Private WithEvents chkFt5 As MSForms.CheckBox
Private WithEvents chkFt6 As MSForms.CheckBox
Private WithEvents chkFt7 As MSForms.CheckBox
Private WithEvents chkFt8 As MSForms.CheckBox
Private WithEvents chkFt9 As MSForms.CheckBox
Private WithEvents chkFt10 As MSForms.CheckBox
Private WithEvents cboExisting As MSForms.ComboBox
Private chkLog As MSForms.CheckBox
Private WithEvents chkAbortEmpty As MSForms.CheckBox
Private WithEvents chkStopEmptyRecord As MSForms.CheckBox
Private cboConfirm As MSForms.ComboBox
Private cboResult As MSForms.ComboBox
Private WithEvents chkQuickScan As MSForms.CheckBox
Private cboOpenFolder As MSForms.ComboBox
Private WithEvents cboAbortField As MSForms.ComboBox
Private fraPreview As MSForms.Frame
Private lblCount As MSForms.Label
Private WithEvents btnPreview As MSForms.CommandButton
Private WithEvents lstPreview As MSForms.ListBox
Private lblDetail As MSForms.Label
Private fraConflict As MSForms.Frame
Private lblConflict As MSForms.Label
Private chkConflictAll As MSForms.CheckBox
Private WithEvents btnConflictOverwrite As MSForms.CommandButton
Private WithEvents btnConflictSkip As MSForms.CommandButton
Private WithEvents btnConflictNumber As MSForms.CommandButton
Private WithEvents btnConflictAbort As MSForms.CommandButton
Private lblProgress As MSForms.Label
Private lblProgressBar As MSForms.Label
Private lblProgressBack As MSForms.Label
Private fraOptions As MSForms.Frame
Private fraProgress As MSForms.Frame
Private fraMessages As MSForms.Frame
Private WithEvents cboLanguage As MSForms.ComboBox
Private lblLanguage As MSForms.Label
Private WithEvents lblOldSite As MSForms.Label
Private lblEta As MSForms.Label
Private lblExisting As MSForms.Label
Private lblConfirm As MSForms.Label
Private lblResultMsg As MSForms.Label
Private lblOpenFolder As MSForms.Label

' width of the progress bar in the current dialog size
Private ProgressBarWidth As Single

' size of the dialog, used to resize the dialog (windows only)
Private ResizeReady As Boolean
Private InResize As Boolean
Private MinInsideWidth As Single
Private MinInsideHeight As Single
Private LastInsideWidth As Single
Private LastInsideHeight As Single

' internal names of the file types (check boxes chkFt1 to chkFt10), in the order the files are saved
Private FiletypeIds() As String
Private FiletypeCount As Long
Private FiletypeLeft As Single
Private FiletypeTop As Single
Private FiletypeWidth As Single

' cached data source values, read once when the form is opened
Private FieldNames() As String
Private FieldCount As Long
Private RecordNumbers() As Long
Private RecordValues() As String
Private CachedRecordCount As Long
Private RecordsInSource As Long
Private ScanStoppedAt As Long
Private CachedWithStop As Boolean
Private CachedAbortIndex As Long
Private CachedQuick As Boolean
Private QuickLimited As Boolean
Private PreviewConflicts As Long
Private CurrentFiles As Long
Private StatusKey As String
Private EtaSum As Double
Private PriorSecondsPerLetter As Double
Private EtaCount As Long
Private ExportStart As Double
Private ScanStart As Double
Private LoadForSave As Boolean
Private CloseAfterLoad As Boolean

' state of the form and of the running export
Private IsInitializing As Boolean
Private IsLoading As Boolean
Private DataLoaded As Boolean
Private IsRunning As Boolean
Private CancelRequested As Boolean
Private ConflictChoice As String
Private LogFileNo As Integer


Private Sub UserForm_Initialize()
    Dim TmpFiletypes As String
    Dim TmpMode As String
    Dim TmpField As String
    Dim TmpAbort As Boolean
    Dim Modes As Variant
    Dim Id As Variant
    Dim i As Long

    IsInitializing = True

    ' replace the single file type choice with a multi selection list
    BuildFiletypeList
    BuildMessagesSection

    ' set default filetypes
    If StorageFiletype <> "" Then
        TmpFiletypes = StorageFiletype
    Else
        TmpFiletypes = DefaultFiletype
    End If

    For Each Id In Split(TmpFiletypes, ";")
        For i = 0 To FiletypeCount - 1
            If StrComp(FiletypeIds(i), Trim(Id), vbTextCompare) = 0 Then GetFiletypeBox(i).Value = True
        Next i
    Next Id
    If GetSelectedFiletypes() = "" Then GetFiletypeBox(0).Value = True

    ' set default filelocation
    If StorageFilelocation <> "" Then
        Me.filelocation.Value = StorageFilelocation
    Else
        If DefaultFilelocation = "desktop" Or DefaultFilelocation = "Desktop" Then
            If System.OperatingSystem Like "*Mac*" Then
                Me.filelocation.Value = MacScript("return (path to desktop folder) as string")
            Else
                Me.filelocation.Value = CreateObject("WScript.Shell").SpecialFolders("Desktop")
            End If
        Else
            Me.filelocation.Value = DefaultFilelocation
        End If
    End If

    If IsNewMac() Then
        ' new office for mac needs dedicated path selection because of access rights (sandbox)
        Me.filelocation.Value = ""
    End If

    ' set default foldername
    If StorageFoldername <> "" Then
        Me.foldername.Value = StorageFoldername
    Else
        Me.foldername.Value = DefaultFoldername
    End If

    ' set default filenamepattern
    If StorageFilenamepattern <> "" Then
        Me.filenamepattern.Value = StorageFilenamepattern
    ElseIf HasCustomDefaultPattern() Then
        ' own default pattern of the user (module formletter2pdf); it is not the shipped pattern, so it passes CheckFilenamePattern
        Me.filenamepattern.Value = Trim(CStr(CustomDefaultFilenamePattern))
    Else
        Me.filenamepattern.Value = DefaultFilenamepattern
    End If

    ' show current version (caption in the form layout still says v2.10) and add it to the title of the window
    On Error Resume Next
    Me.Controls("label_version").Caption = "v" & AppVersion & " |"
    On Error GoTo 0
    Me.Caption = "formletter2pdf v" & AppVersion

    ' read the field names and create the additional controls; the records are read by LoadRecords
    ReadFieldNames
    BuildExtendedControls

    ' set default handling of existing files
    If StorageExistingFiles <> "" Then
        TmpMode = StorageExistingFiles
    Else
        TmpMode = DefaultExistingFiles
    End If
    cboExisting.ListIndex = 0
    Modes = Split(ExistingModes, ",")
    For i = 0 To UBound(Modes)
        If LCase(TmpMode) = Modes(i) Then cboExisting.ListIndex = i
    Next i

    ' set default log option
    If StorageWriteLog <> "" Then
        chkLog.Value = (StorageWriteLog = "1")
    Else
        chkLog.Value = DefaultWriteLog
    End If

    ' set default field which must not be empty
    If StorageAbortIfEmpty <> "" Then
        TmpAbort = (StorageAbortIfEmpty = "1")
        TmpField = StorageAbortIfEmptyField
    Else
        TmpAbort = (DefaultAbortIfEmptyField <> "")
        TmpField = DefaultAbortIfEmptyField
    End If
    If TmpField <> "" Then
        If Not SelectListItem(cboAbortField, TmpField) Then TmpAbort = False
    End If
    chkAbortEmpty.Value = TmpAbort

    ' set default stop at the first record where all fields are empty
    If StorageStopAtEmptyRecord <> "" Then
        chkStopEmptyRecord.Value = (StorageStopAtEmptyRecord = "1")
    Else
        chkStopEmptyRecord.Value = DefaultStopAtEmptyRecord
    End If

    ' all texts of the dialog in the language of the user
    CurrentFiles = -1
    StatusKey = "S_SCAN_HINT"
    ApplyLanguage

    ' the donation text is cut off at the bottom left: let it wrap into two lines next to the save button and make room for them
    On Error Resume Next
    With Me.Controls("label_donate")
        .WordWrap = True
        .Width = Me.Controls("create").Left - .Left - 8
        If .Height < 28 Then
            Me.Height = Me.Height + (28 - .Height)
            .Height = 28
        End If
    End With
    On Error GoTo 0

    IsInitializing = False

End Sub


Private Sub ApplyLanguage()
    ' sets all texts of the dialog in the current language; also used when the language is changed in the dialog
    Dim WasInitializing As Boolean

    WasInitializing = IsInitializing
    IsInitializing = True

    ' original controls: their texts are stored in english in the form layout (.frx), so they are set here
    On Error Resume Next
    Me.Controls("label_filetype").Caption = Tx("L_FILETYPE")
    Me.Controls("label_subheadline").Caption = Tx("L_SUBHEAD")
    PlaceOldSiteLink
    Me.Controls("label_foldername").Caption = Tx("L_FOLDER")
    Me.Controls("label_foldername_hint").Caption = Tx("L_FOLDER_HINT")
    Me.Controls("label_filename").Caption = Tx("L_FILE")
    Me.Controls("label_filename_hint").Caption = Tx("L_FILE_HINT")
    Me.Controls("label_filename_hint_action").Caption = Tx("L_FILE_HINT_ACTION")
    Me.Controls("label_filelocation").Caption = Tx("L_SAVETO")
    Me.Controls("label_filelocation_help").Caption = Tx("L_SAVETO_HELP")
    Me.Controls("filelocation_choose").Caption = Tx("B_CHOOSE")
    Me.Controls("option_print").Caption = Tx("L_PRINT")
    Me.Controls("label_option_print").Caption = Tx("L_PRINT_HINT")
    Me.Controls("label_donate").Caption = Tx("L_DONATE")
    Me.Controls("cancel").Caption = Tx("B_CLOSE")
    On Error GoTo 0

    ' additional controls
    lblLanguage.Caption = Tx("L_LANGUAGE")
    fraOptions.Caption = Tx("F_OPTIONS")
    lblExisting.Caption = Tx("L_EXISTING")
    FillCombo cboExisting, "CBO_EX_ASK,CBO_EX_OVERWRITE,CBO_EX_SKIP,CBO_EX_NUMBER"
    chkLog.Caption = TxF("O_LOG", LogFileName)
    chkAbortEmpty.Caption = Tx("O_STOP_FIELD")
    chkStopEmptyRecord.Caption = Tx("O_STOP_ALL")
    fraPreview.Caption = Tx("F_PREVIEW")
    chkQuickScan.Caption = TxF("O_QUICK", QuickScanLimit)
    fraConflict.Caption = Tx("F_CONFLICT")
    btnConflictOverwrite.Caption = Tx("CBO_EX_OVERWRITE")
    btnConflictSkip.Caption = Tx("CBO_EX_SKIP")
    btnConflictNumber.Caption = Tx("CBO_EX_NUMBER")
    btnConflictAbort.Caption = Tx("B_ABORT")
    chkConflictAll.Caption = Tx("O_APPLY_ALL")
    fraProgress.Caption = Tx("F_PROGRESS")
    fraMessages.Caption = Tx("F_MESSAGES")
    lblConfirm.Caption = Tx("L_MSG_CONFIRM")
    FillCombo cboConfirm, "M_ALWAYS,M_WARNINGS,M_NEVER"
    lblResultMsg.Caption = Tx("L_MSG_RESULT")
    FillCombo cboResult, "M_ALWAYS,M_WARNINGS,M_NEVER"
    If Not cboOpenFolder Is Nothing Then
        lblOpenFolder.Caption = Tx("L_MSG_OPEN")
        FillCombo cboOpenFolder, "M_ASK,M_OPEN,M_NOOPEN"
    End If

    IsInitializing = WasInitializing

    ' texts which depend on the state of the dialog
    UpdatePreviewButton
    SetSaveButton IsRunning Or (IsLoading And LoadForSave)
    If Not IsRunning And Not IsLoading Then lblProgress.Caption = Tx("PR_NOT_STARTED")
    If Not IsLoading Then
        If DataLoaded Then
            RefreshPreview
        Else
            lblCount.ForeColor = vbButtonText
            lblCount.Caption = TxF(StatusKey, Tx("B_SCAN"))
        End If
    End If
    ShowPreviewDetail
End Sub


Private Sub FillCombo(Box As MSForms.ComboBox, ByVal Keys As String)
    ' fills a drop-down list with the texts of the keys (separated by ","), the selected entry stays selected
    Dim Index As Long
    Dim Key As Variant

    Index = Box.ListIndex
    Box.Clear
    For Each Key In Split(Keys, ",")
        Box.AddItem Tx(CStr(Key))
    Next Key
    If Index < 0 Then Index = 0
    Box.ListIndex = Index
End Sub


Private Sub cboLanguage_Change()
    ' switches all texts of the dialog to the chosen language
    If IsInitializing Or IsRunning Or IsLoading Then Exit Sub

    SetLanguage IIf(cboLanguage.ListIndex = 1, LangGerman, LangEnglish)
    StorageLanguage = CurrentLanguage
    ApplyLanguage
End Sub


Private Sub UserForm_QueryClose(Cancel As Integer, CloseMode As Integer)
    ' closing the window while reading the data source or during the export cancels it instead
    If IsRunning Or IsLoading Then
        Cancel = True
        ' closing the dialog while the data source is read for the preview closes it after the abort
        If IsLoading And Not LoadForSave Then CloseAfterLoad = True
        RequestCancel
    ElseIf ResizeReady And Not IsMac() Then
        ' remember the size of the dialog for the next start (until word is closed)
        StorageFormWidth = Me.Width
        StorageFormHeight = Me.Height
    End If
End Sub


Private Sub UserForm_Activate()
    ' make the dialog resizable once (windows only)
    Dim BeforeWidth As Single
    Dim BeforeHeight As Single

    If ResizeReady Then Exit Sub

    If Not IsMac() Then
        ' the resize border takes space from the client area, which would cut off the controls at the right and bottom border;
        ' the dialog is made bigger by that space and some more
        BeforeWidth = Me.InsideWidth
        BeforeHeight = Me.InsideHeight
        EnableResize
        If Me.InsideWidth < BeforeWidth Then Me.Width = Me.Width + (BeforeWidth - Me.InsideWidth)
        If Me.InsideHeight < BeforeHeight Then Me.Height = Me.Height + (BeforeHeight - Me.InsideHeight)
        Me.Height = Me.Height + 8
    End If
    MinInsideWidth = Me.InsideWidth
    MinInsideHeight = Me.InsideHeight
    LastInsideWidth = MinInsideWidth
    LastInsideHeight = MinInsideHeight
    ResizeReady = True

    ' restore the remembered size; the layout follows in UserForm_Resize, smaller than the start size is not possible
    If Not IsMac() And StorageFormWidth > 0 And StorageFormHeight > 0 Then
        On Error Resume Next
        If StorageFormWidth > Me.Width Then Me.Width = StorageFormWidth
        If StorageFormHeight > Me.Height Then Me.Height = StorageFormHeight
        On Error GoTo 0
    End If
End Sub


Private Sub EnableResize()
    ' vba user forms have no resize border; it is added with the window style of the dialog
#If Mac Then
    ' not available on mac
#Else
    Const GWL_STYLE As Long = -16
    Const WS_THICKFRAME As Long = &H40000
    Const WS_MAXIMIZEBOX As Long = &H10000

    #If VBA7 Then
        Dim hWnd As LongPtr
    #Else
        Dim hWnd As Long
    #End If
    Dim OriginalCaption As String
    Dim Style As Long

    On Error Resume Next

    ' find the window of this dialog by a temporary unique caption
    OriginalCaption = Me.Caption
    Me.Caption = "formletter2pdf " & Format(Now, "yyyymmddhhnnss") & CStr(Timer)
    hWnd = FindWindow("ThunderDFrame", Me.Caption)
    Me.Caption = OriginalCaption
    If hWnd = 0 Then Exit Sub

    Style = GetWindowLong(hWnd, GWL_STYLE)
    SetWindowLong hWnd, GWL_STYLE, Style Or WS_THICKFRAME Or WS_MAXIMIZEBOX
    DrawMenuBar hWnd

    On Error GoTo 0
#End If
End Sub


Private Sub UserForm_Resize()
    Dim NewWidth As Single
    Dim NewHeight As Single

    If Not ResizeReady Or InResize Then Exit Sub
    InResize = True
    On Error Resume Next

    ' the dialog must not get smaller than it was when it was opened
    If Me.InsideWidth < MinInsideWidth Then Me.Width = Me.Width + (MinInsideWidth - Me.InsideWidth)
    If Me.InsideHeight < MinInsideHeight Then Me.Height = Me.Height + (MinInsideHeight - Me.InsideHeight)
    NewWidth = Me.InsideWidth
    NewHeight = Me.InsideHeight

    If NewWidth <> LastInsideWidth Or NewHeight <> LastInsideHeight Then
        ApplyLayout NewWidth - LastInsideWidth, NewHeight - LastInsideHeight
        LastInsideWidth = NewWidth
        LastInsideHeight = NewHeight
    End If

    On Error GoTo 0
    InResize = False
End Sub


Private Sub ApplyLayout(ByVal DeltaWidth As Single, ByVal DeltaHeight As Single)
    ' the original controls (left) keep their size; the additional panel (right) grows with the dialog:
    ' wider panel and bigger file name list, progress at the bottom
    fraOptions.Width = fraOptions.Width + DeltaWidth

    fraPreview.Width = fraPreview.Width + DeltaWidth
    fraPreview.Height = fraPreview.Height + DeltaHeight
    lblCount.Width = lblCount.Width + DeltaWidth
    btnPreview.Left = btnPreview.Left + DeltaWidth
    lstPreview.Width = lstPreview.Width + DeltaWidth
    lstPreview.Height = lstPreview.Height + DeltaHeight
    lblDetail.Top = lblDetail.Top + DeltaHeight
    lblDetail.Width = lblDetail.Width + DeltaWidth

    fraConflict.Width = fraConflict.Width + DeltaWidth
    lblConflict.Width = lblConflict.Width + DeltaWidth

    fraProgress.Top = fraProgress.Top + DeltaHeight
    fraProgress.Width = fraProgress.Width + DeltaWidth
    lblProgress.Width = lblProgress.Width + DeltaWidth
    lblEta.Width = lblEta.Width + DeltaWidth
    lblProgressBack.Width = lblProgressBack.Width + DeltaWidth
    ProgressBarWidth = ProgressBarWidth + DeltaWidth
End Sub


Private Sub chkFt1_Click()
    RefreshPreview
End Sub


Private Sub chkFt2_Click()
    RefreshPreview
End Sub


Private Sub chkFt3_Click()
    RefreshPreview
End Sub


Private Sub chkFt4_Click()
    RefreshPreview
End Sub


Private Sub chkFt5_Click()
    RefreshPreview
End Sub


Private Sub chkFt6_Click()
    RefreshPreview
End Sub


Private Sub chkFt7_Click()
    RefreshPreview
End Sub


Private Sub chkFt8_Click()
    RefreshPreview
End Sub


Private Sub chkFt9_Click()
    RefreshPreview
End Sub


Private Sub chkFt10_Click()
    RefreshPreview
End Sub


Private Sub filelocation_AfterUpdate()
    RefreshPreview
End Sub


Private Sub foldername_AfterUpdate()
    RefreshPreview
End Sub


Private Sub filenamepattern_AfterUpdate()
    RefreshPreview
End Sub


Private Sub cboExisting_Change()
    RefreshPreview
End Sub


Private Sub chkQuickScan_Click()
    RefreshOrInvalidate
End Sub


Private Sub chkStopEmptyRecord_Click()
    RefreshOrInvalidate
End Sub


Private Sub SetSaveButton(ByVal Aborting As Boolean)
    ' the save button shows the number of files to be saved ("Save 12 files") as soon as it is known (after a scan)
    ' and becomes "Abort export" while the export runs (or the data source is read for it)
    If Aborting Then
        Me.Controls("create").Caption = Tx("B_ABORT_SAVE")
    ElseIf CurrentFiles >= 0 Then
        Me.Controls("create").Caption = TxF("B_SAVE_N", TxC(CurrentFiles, "N_FILE"))
    Else
        Me.Controls("create").Caption = Tx("B_SAVE_DEFAULT")
    End If
End Sub


Private Sub SetCreateCaption(ByVal Files As Long)
    ' remembers the number of files to be saved (-1: not known) and shows it on the save button
    CurrentFiles = Files
    If Not IsRunning And Not (IsLoading And LoadForSave) Then SetSaveButton False
End Sub


Private Sub UpdatePreviewButton()
    ' the scan button reads the data source: "Scan" as long as no records are read, afterwards "Refresh", "Abort scan" while reading
    If btnPreview Is Nothing Then Exit Sub
    If IsLoading Then
        btnPreview.Caption = Tx("B_ABORT_SCAN")
    ElseIf DataLoaded Then
        btnPreview.Caption = Tx("B_REFRESH")
    Else
        btnPreview.Caption = Tx("B_SCAN")
    End If
End Sub


Private Sub SetStatus(ByVal Key As String)
    ' status text above the preview list while no preview is shown (the key is remembered for a change of the language)
    StatusKey = Key
    lblCount.ForeColor = vbButtonText
    lblCount.Caption = TxF(Key, Tx("B_SCAN"))
End Sub


Private Sub RefreshOrInvalidate()
    ' the stop options can cut the records short while reading; if an option changes so that records
    ' which were not read are needed, the records have to be read again
    If IsInitializing Or IsRunning Or IsLoading Then Exit Sub

    If DataLoaded And (ScanStoppedAt > 0 Or QuickLimited) And Not IsCacheStillValid() Then
        DataLoaded = False
        UpdatePreviewButton
        SetCreateCaption -1
        lstPreview.Clear
        ShowPreviewDetail
        SetStatus "S_SCAN_AGAIN"
    Else
        RefreshPreview
    End If
End Sub


Private Function IsCacheStillValid() As Boolean
    IsCacheStillValid = True
    If CachedWithStop And chkStopEmptyRecord.Value <> True Then IsCacheStillValid = False
    If CachedAbortIndex > 0 And GetAbortFieldIndex() <> CachedAbortIndex Then IsCacheStillValid = False
    If CachedQuick And chkQuickScan.Value <> True Then IsCacheStillValid = False
End Function


Private Sub chkAbortEmpty_Click()
    RefreshOrInvalidate
End Sub


Private Sub cboAbortField_Change()
    ' choosing a field activates the option
    If Not IsInitializing And cboAbortField.ListIndex >= 0 And chkAbortEmpty.Value = False Then
        chkAbortEmpty.Value = True
    Else
        RefreshOrInvalidate
    End If
End Sub


Private Sub lstPreview_Change()
    ShowPreviewDetail
End Sub


Private Sub ShowPreviewDetail()
    ' shows the complete text of the selected row, because long file names are cut off in the list
    Dim Row As Long
    Dim Text As String

    Row = lstPreview.ListIndex
    If Row < 0 Then
        lblDetail.Caption = Tx("P_SELECT_ROW")
        Exit Sub
    End If

    If IsNumeric(lstPreview.List(Row, 0)) Then
        Text = TxF("P_RECORD", lstPreview.List(Row, 0))
        If lstPreview.List(Row, 1) <> "" Then Text = Text & ": " & lstPreview.List(Row, 1)
    Else
        ' information row (e.g. quick scan hint) without a record number
        Text = lstPreview.List(Row, 1)
    End If
    If lstPreview.List(Row, 2) <> "" Then Text = Text & vbNewLine & lstPreview.List(Row, 2)
    lblDetail.Caption = Text
End Sub

Private Sub btnPreview_Click()
    ' the data source is only read when needed, because this can take a long time
    If IsRunning Then Exit Sub

    ' while the data source is read, the button is the "Abort scan" button
    If IsLoading Then
        RequestCancel
        Exit Sub
    End If

    ' check the settings before the data source is read, because reading can take long
    If Not ValidateInput(False) Then Exit Sub

    If Not DataLoaded Then LoadRecords (chkQuickScan.Value = True)
    CancelRequested = False
    If CloseAfterLoad Then
        Unload Me
        Exit Sub
    End If
    If DataLoaded Then
        RefreshPreview
    Else
        SetStatus "S_SCAN_ABORTED"
    End If
End Sub


Private Function HasCustomDefaultPattern() As Boolean
    ' True if CustomDefaultFilenamePattern (module formletter2pdf) is neither Null nor empty
    If IsNull(CustomDefaultFilenamePattern) Then Exit Function
    HasCustomDefaultPattern = (Trim(CustomDefaultFilenamePattern & "") <> "")
End Function


Private Function CheckFilenamePattern() As Boolean
    ' the file name pattern which is shipped as default value of formletter2pdf uses fields (last_name, ID) which
    ' normally do not exist in a data source, so it has to be changed before scanning or exporting.
    ' The pattern is compared with the hard coded value ShippedFilenamepattern on purpose, NOT with
    ' DefaultFilenamepattern: this way the default value can be changed in the module formletter2pdf to
    ' an own pattern, which is then accepted without this error.
    If StrComp(Trim(Me.filenamepattern.Value), ShippedFilenamepattern, vbTextCompare) = 0 Then
        MsgBox TxF("E_SHIPPED", ShippedFilenamepattern), vbOKOnly + vbCritical
        On Error Resume Next
        Me.filenamepattern.SetFocus
        On Error GoTo 0
        Exit Function
    End If
    CheckFilenamePattern = True
End Function


Private Function ValidateInput(ByVal ForExport As Boolean) As Boolean
    ' checks the settings of the dialog before the data source is read (scan) or any file is written (export),
    ' because both can take long; the checks for the target folder are only needed for the export
    Dim Duplicates As String
    Dim fso As Object

    ' the file name pattern which is shipped with formletter2pdf has to be changed
    If Not CheckFilenamePattern() Then Exit Function

    ' file types
    If GetSelectedFiletypes() = "" Then
        MsgBox Tx("E_NO_FILETYPE"), vbOKOnly + vbCritical
        Exit Function
    End If
    Duplicates = FindDuplicateFiletypes()
    If Duplicates <> "" Then
        MsgBox TxF("E_DUP_TYPES", Duplicates), vbOKOnly + vbCritical
        Exit Function
    End If

    ' file name pattern
    If Trim(Me.filenamepattern.Value) = "" Then
        MsgBox Tx("E_NO_PATTERN"), vbOKOnly + vbCritical
        Exit Function
    End If

    ' field which must not be empty
    If chkAbortEmpty.Value = True And cboAbortField.ListIndex < 0 Then
        MsgBox Tx("E_NO_FIELD"), vbOKOnly + vbCritical
        Exit Function
    End If

    If ForExport Then
        ' storage location
        If Me.filelocation.Value = "" Then
            MsgBox TxF("E_NO_LOCATION", Tx("L_SAVETO")), vbOKOnly + vbCritical
            Exit Function
        End If
        If Not IsMac() Then
            On Error Resume Next
            Set fso = CreateObject("Scripting.FileSystemObject")
            On Error GoTo 0
            If Not fso Is Nothing Then
                If Not fso.FolderExists(CleanLocation(Me.filelocation.Value)) Then
                    MsgBox TxF("E_LOC_MISSING", Me.filelocation.Value, Tx("L_SAVETO")), vbOKOnly + vbCritical
                    Exit Function
                End If
                If Not CanWriteTo(CleanLocation(Me.filelocation.Value)) Then
                    MsgBox TxF("E_LOC_WRITE", Me.filelocation.Value, Tx("L_SAVETO")), vbOKOnly + vbCritical
                    Exit Function
                End If
            End If
        End If

        ' folder name
        If Trim(Me.foldername.Value) = "" Then
            MsgBox Tx("E_NO_FOLDER"), vbOKOnly + vbCritical
            Exit Function
        ElseIf HasInvalidCharacters(Me.foldername.Value) Then
            MsgBox TxF("E_FOLDER_CHARS", "\ / : * ? " & Chr(34) & " < > |"), vbOKOnly + vbCritical
            Exit Function
        End If
    End If

    ValidateInput = True
End Function

Private Function CanWriteTo(ByVal FolderPath As String) As Boolean
    ' tests the write access by creating and deleting a small test file in the folder
    Dim TestFile As String
    Dim FileNo As Integer

    TestFile = FolderPath & "\formletter2pdf-" & Format(Now, "yyyymmddhhnnss") & ".tmp"
    On Error GoTo Failed
    FileNo = FreeFile
    Open TestFile For Output As #FileNo
    Close #FileNo
    Kill TestFile
    CanWriteTo = True
    Exit Function

Failed:
    On Error Resume Next
    Close #FileNo
    Kill TestFile
End Function


Private Function IsReservedName(ByVal FileBase As String) As Boolean
    ' windows does not allow files with the names CON, PRN, AUX, NUL, COM1-9 and LPT1-9, also not with an extension
    Dim Stem As String
    Dim p As Long

    If IsMac() Then Exit Function

    Stem = FileBase
    p = InStr(Stem, ".")
    If p > 0 Then Stem = Left(Stem, p - 1)
    Stem = UCase(Trim(Stem))

    Select Case Stem
    Case "CON", "PRN", "AUX", "NUL"
        IsReservedName = True
    Case Else
        If Len(Stem) = 4 Then
            If Left(Stem, 3) = "COM" Or Left(Stem, 3) = "LPT" Then
                IsReservedName = (Mid(Stem, 4, 1) Like "[1-9]")
            End If
        End If
    End Select
End Function


Private Function CheckFileNames(ByVal SavePath As String, Extensions() As String, ByVal Pattern As String, ByVal EmptyStopIndex As Long, ByVal StopRecord As Long) As Boolean
    ' checks the file names of all form letters before any file is written (windows only): names reserved by windows
    ' and paths which are too long would make saving fail in the middle of the export
    Dim Documentname As String
    Dim BaseName As String
    Dim ReservedList As String
    Dim LongestName As String
    Dim LongestExt As Long
    Dim Reserve As Long
    Dim PathLen As Long
    Dim LongestLen As Long
    Dim LongestRecord As Long
    Dim TooLongCount As Long
    Dim ReservedCount As Long
    Dim Counter As Long
    Dim i As Long
    Dim t As Long

    CheckFileNames = True
    If IsMac() Then Exit Function

    Documentname = GetDocumentName(ActiveDocument)
    For t = LBound(Extensions) To UBound(Extensions)
        If Len(Extensions(t)) > LongestExt Then LongestExt = Len(Extensions(t))
    Next t

    ' if numbers are added to existing names, they need some more characters
    Select Case GetExistingMode()
    Case "overwrite", "skip"
        Reserve = 0
    Case Else
        Reserve = PathNumberReserve
    End Select

    For i = 1 To CachedRecordCount
        If EmptyStopIndex > 0 And i >= EmptyStopIndex Then Exit For
        If StopRecord > 0 And RecordNumbers(i) = StopRecord Then Exit For
        If Not IsRecordEmpty(i) Then
            Counter = Counter + 1
            BaseName = BuildFileBaseName(i, Counter, Pattern, Documentname)

            If IsReservedName(BaseName) Then
                ReservedCount = ReservedCount + 1
                If ReservedCount <= 5 Then ReservedList = ReservedList & TxF("P_RECORD", RecordNumbers(i)) & ": " & BaseName & vbNewLine
            End If

            PathLen = Len(SavePath) + Len(BaseName) + 1 + LongestExt + Reserve
            If PathLen > MaxPathLength Then
                TooLongCount = TooLongCount + 1
                If PathLen > LongestLen Then
                    LongestLen = PathLen
                    LongestRecord = RecordNumbers(i)
                    LongestName = BaseName
                End If
            End If
        End If
    Next i

    If ReservedCount > 0 Then
        MsgBox TxC(ReservedCount, "E_RESERVED") & vbNewLine & vbNewLine & ReservedList & _
            IIf(ReservedCount > 5, "..." & vbNewLine, "") & vbNewLine & _
            Tx("E_RESERVED_HINT"), vbOKOnly + vbCritical
        CheckFileNames = False
    ElseIf TooLongCount > 0 Then
        MsgBox TxC(TooLongCount, "E_PATH", MaxPathLength) & vbNewLine & vbNewLine & _
            TxF("E_PATH_INFO", LongestLen, LongestRecord, LongestName), vbOKOnly + vbCritical
        CheckFileNames = False
    End If
End Function


Private Function FindDuplicateFiletypes() As String
    ' lists the selected file types which share a file extension, one line per extension, e.g.
    ' ".pdf: Portable Document Format (.pdf), PDF/A (ISO 19005-1, .pdf)"; empty if there is none
    Dim Done As New Collection
    Dim Ext As String
    Dim Captions As String
    Dim Count As Long
    Dim i As Long
    Dim j As Long

    For i = 0 To FiletypeCount - 1
        If GetFiletypeBox(i).Value = True Then
            Ext = GetExtension(FiletypeIds(i))
            If AddToSet(Done, Ext) Then
                Captions = ""
                Count = 0
                For j = 0 To FiletypeCount - 1
                    If GetFiletypeBox(j).Value = True Then
                        If GetExtension(FiletypeIds(j)) = Ext Then
                            Count = Count + 1
                            If Captions <> "" Then Captions = Captions & ", "
                            Captions = Captions & GetFiletypeBox(j).Caption
                        End If
                    End If
                Next j
                If Count > 1 Then FindDuplicateFiletypes = FindDuplicateFiletypes & "." & Ext & ": " & Captions & vbNewLine
            End If
        End If
    Next i
End Function


Private Sub btnConflictOverwrite_Click()
    ConflictChoice = "overwrite"
End Sub


Private Sub btnConflictSkip_Click()
    ConflictChoice = "skip"
End Sub


Private Sub btnConflictNumber_Click()
    ConflictChoice = "number"
End Sub


Private Sub btnConflictAbort_Click()
    ConflictChoice = "abort"
End Sub


Private Sub label_filename_hint_action_Click()
    Dim Message As String
    Dim i As Long

    For i = 1 To FieldCount
        Message = Message & "{" & FieldNames(i) & "}" & vbCrLf
    Next i

    MsgBox Message, vbOKOnly + vbInformation
End Sub


Private Sub FileLocation_choose_Click()

    If System.OperatingSystem Like "*Mac*" Then

            'define Variables
            Dim MacScriptPath As String

            ' create browse for folder dialog
            On Error Resume Next

            If IsNewMac() Then
                MacScriptPath = MacScript("(choose folder with prompt """ & Tx("L_SAVETO_HELP") & """" & _
    "default location alias """ & MacScript("return (path to desktop folder) as String") & """) as string")
            Else
                MacScriptPath = MacScript("(choose folder with prompt """ & Tx("L_SAVETO_HELP") & """" & _
    "default location alias """ & Me.filelocation.Value & """) as string")
            End If

            On Error GoTo 0

            ' save choosen path to FileLocation
            If MacScriptPath <> "" Then
                Me.filelocation.Value = MacScriptPath
            End If

    Else

        ' check initial path
        Dim tmpPath As String
        Dim PathWithTrailingSlash As String
        tmpPath = Me.filelocation.Value
        Select Case Right(tmpPath, 1)
        Case "\"
            PathWithTrailingSlash = tmpPath
        Case Else
            PathWithTrailingSlash = tmpPath & "\"
        End Select

        ' create browse for folder dialog
        With Application.FileDialog(msoFileDialogFolderPicker)
            .InitialFileName = PathWithTrailingSlash
            .Title = Tx("L_SAVETO_HELP")
            .AllowMultiSelect = False
            If .Show = -1 Then
                Me.filelocation.Value = .SelectedItems(1)
            End If
        End With

    End If

    RefreshPreview

End Sub


Private Sub cancel_Click()

    If IsRunning Then
        ' close while the export runs: stop the export after the current form letter, the dialog closes afterwards
        RequestCancel
    ElseIf IsLoading Then
        ' stop reading the data source; "Close" also closes the dialog afterwards, "Abort export" (read for the export) does not
        If Not LoadForSave Then CloseAfterLoad = True
        RequestCancel
    Else
        ' "Close": quit macro
        Unload Me
    End If

End Sub


Private Sub create_Click()

    '
    ' define variables
    '
    Dim Filetypes As Variant
    Dim Extensions() As String
    Dim Targets() As String
    Dim Actions() As String
    Dim Location As String
    Dim Folder As String
    Dim Pattern As String
    Dim SavePath As String
    Dim DirCheckPath As String
    Dim Documentname As String
    Dim ExistingMode As String
    Dim Mode As String
    Dim BaseName As String
    Dim TargetName As String
    Dim ExistingNames As String
    Dim MsgText As String
    Dim ResultText As String
    Dim ErrDescription As String
    Dim isDir As Boolean
    Dim AnyToSave As Boolean
    Dim PrintFiles As Boolean
    Dim PrintDialogShown As Boolean
    Dim LogWritten As Boolean
    Dim Stopped As Boolean
    Dim HasWarnings As Boolean
    Dim ShowResult As Boolean
    Dim ConfirmMode As String
    Dim ResultMode As String
    Dim QuestionText As String
    Dim StoppedEmpty As Boolean
    Dim EmptyStopIndex As Long
    Dim TypeCount As Long
    Dim AbortIndex As Long
    Dim StopRecord As Long
    Dim i As Long
    Dim t As Long
    Dim Counter As Long
    Dim TotalCount As Long
    Dim PlannedCount As Long
    Dim EmptyCount As Long
    Dim LetterCount As Long
    Dim FileCount As Long
    Dim OverwrittenCount As Long
    Dim NumberedCount As Long
    Dim SkippedCount As Long
    Dim DocumentCount As Long
    Dim CurrentRecord As Long
    Dim ErrNumber As Long
    Dim SavedDestination As Long
    Dim SavedFirstRecord As Long
    Dim SavedLastRecord As Long
    Dim SavedActiveRecord As Long
    Dim MainDocument As Document
    Dim OutputDocument As Document
    Dim LetterStart As Double
    Dim SkipSample As Boolean

    ' while the export runs (or the data source is read for it) the button is the "Abort export" button
    If IsRunning Or (IsLoading And LoadForSave) Then
        RequestCancel
        Exit Sub
    End If
    If IsLoading Then Exit Sub

    ' check all settings before the data source is read and before any file is written, because both can take long
    If Not ValidateInput(True) Then Exit Sub

    ' the records are needed for the export; if the preview was not used yet or only a quick scan was done, read them now
    If Not DataLoaded Or QuickLimited Then
        btnPreview.Enabled = False
        SetSaveButton True
        LoadForSave = True
        LoadRecords False
        LoadForSave = False
        SetSaveButton False
        btnPreview.Enabled = True
    End If
    CancelRequested = False
    If Not DataLoaded Then
        SetStatus "S_SCAN_ABORTED"
        Exit Sub
    End If


    '
    ' get values from form
    '

        ' the values have been checked in ValidateInput
    Filetypes = Split(GetSelectedFiletypes(), ";")
    TypeCount = UBound(Filetypes) + 1
    Extensions = GetExtensions(Filetypes)
    Location = CleanLocation(Me.filelocation.Value)
    Folder = Me.foldername.Value
    Pattern = Me.filenamepattern.Value
    AbortIndex = GetAbortFieldIndex()

    '
    ' preserve values from form for the next run of the macro
    '

    StorageFiletype = GetSelectedFiletypes()
    StorageFilelocation = Me.filelocation.Value
    StorageFoldername = Folder
    StorageFilenamepattern = Pattern
    StorageExistingFiles = GetExistingMode()
    StorageWriteLog = IIf(chkLog.Value = True, "1", "0")
    StorageAbortIfEmpty = IIf(chkAbortEmpty.Value = True, "1", "0")
    StorageAbortIfEmptyField = cboAbortField.Text
    StorageStopAtEmptyRecord = IIf(chkStopEmptyRecord.Value = True, "1", "0")
    If Not IsMac() Then StorageOpenFolder = GetOpenFolderMode()
    StorageConfirmBeforeWriting = GetConfirmMode()
    StorageResultMessage = GetResultMode()
    StorageQuickScan = IIf(chkQuickScan.Value = True, "1", "0")


    '
    ' count the form letters and find the first record with an empty required field
    '

    RefreshPreview
    AbortIndex = GetAbortFieldIndex()
    EmptyStopIndex = GetEmptyStopIndex()

    For i = 1 To CachedRecordCount
        If EmptyStopIndex > 0 And i >= EmptyStopIndex Then Exit For
        If Not IsRecordEmpty(i) Then
            TotalCount = TotalCount + 1
            If StopRecord = 0 Then
                If AbortIndex > 0 Then
                    If Trim(RecordValues(i, AbortIndex)) = "" Then StopRecord = RecordNumbers(i)
                End If
                If StopRecord = 0 Then PlannedCount = PlannedCount + 1
            End If
        End If
    Next i

    If PlannedCount = 0 Then
        If StopRecord > 0 Then
            MsgBox TxF("E_NOTHING_FIELD", FieldNames(AbortIndex), StopRecord), vbOKOnly + vbCritical
        Else
            MsgBox Tx("E_NO_RECORDS"), vbOKOnly + vbCritical
        End If
        Exit Sub
    End If


    '
    ' determine path, ask before start and create target folder
    '

    Folder = Replace(Folder, "{%timestamp%}", Format(Now, TimestampFormat))
    GetTargetPaths Location, Folder, SavePath, DirCheckPath

    ' check the file names of all form letters before any file is written (path length, names reserved by windows)
    If Not CheckFileNames(SavePath, Extensions, Pattern, EmptyStopIndex, StopRecord) Then Exit Sub

    On Error GoTo FolderError

        ' check if target folder already exists
    isDir = TargetFolderExists(SavePath, DirCheckPath)
    If isDir Then
        If MsgBox(TxF("Q_FOLDER_EXISTS", SavePath, Tx("L_EXISTING"), cboExisting.Text), vbYesNo + vbQuestion + vbDefaultButton2) = vbNo Then Exit Sub
    End If

        ' show info dialog
    MsgText = ""
    If StopRecord > 0 Then
        If ScanStoppedAt = StopRecord Then
            ' the data source was not read beyond this record, so the total number of form letters is not known
            MsgText = TxF("C_ATTN_A", FieldNames(AbortIndex), RecordRef(StopRecord)) & vbNewLine & vbNewLine
        Else
            MsgText = TxF("C_ATTN_B", FieldNames(AbortIndex), RecordRef(StopRecord), PlannedCount, TxC(TotalCount, "N_LETTER")) & vbNewLine & vbNewLine
        End If
    End If
    If PreviewConflicts > 0 Then
        MsgText = MsgText & TxC(PreviewConflicts, "C_CONFLICTS", cboExisting.Text) & vbNewLine & vbNewLine
    End If
    If EmptyStopIndex > 0 Then
        MsgText = MsgText & TxF("C_EMPTYSTOP", RecordRef(RecordNumbers(EmptyStopIndex))) & vbNewLine & vbNewLine
    End If
    MsgText = MsgText & TxF("C_SUMMARY", TxC(PlannedCount, "N_LETTER"), TxC(PlannedCount * TypeCount, "N_FILE"), SavePath) & vbNewLine & vbNewLine
    If IsNewMac() Then
        MsgText = MsgText & Tx("C_MAC_NEW")
    ElseIf IsMac() Then
        MsgText = MsgText & Tx("C_MAC_OLD")
    Else
        MsgText = MsgText & Tx("C_WIN")
    End If

        ' the confirm dialog depends on the setting: always, only if there is something to warn about, or never
    HasWarnings = (StopRecord > 0 Or PreviewConflicts > 0 Or FindUnknownMarker(Pattern) <> "")
    ConfirmMode = GetConfirmMode()
    If ConfirmMode = "always" Or (ConfirmMode = "warnings" And HasWarnings) Then
        If MsgBox(MsgText & vbNewLine & vbNewLine & Tx("C_CONTINUE"), vbYesNo + IIf(StopRecord > 0, vbExclamation, vbQuestion) + vbDefaultButton1) = vbNo Then Exit Sub
    End If

        ' create target folder
    If Not isDir Then CreateTargetFolder Location, Folder, SavePath

    On Error GoTo 0


    '
    ' prepare export
    '

    Set MainDocument = ActiveDocument
    Documentname = GetDocumentName(MainDocument)
    ExistingMode = GetExistingMode()
    PrintFiles = (Me.option_print.Value = True)
    ReDim Targets(0 To TypeCount - 1)
    ReDim Actions(0 To TypeCount - 1)

    On Error GoTo ErrorHandling

        ' remember mail merge settings to restore them afterwards
    With MainDocument.MailMerge
        SavedDestination = .Destination
        SavedFirstRecord = .DataSource.FirstRecord
        SavedLastRecord = .DataSource.LastRecord
        SavedActiveRecord = .DataSource.ActiveRecord
    End With

    StartRunning

    If chkLog.Value = True Then
        OpenLog SavePath & LogFileName
        WriteLog "formletter2pdf v" & AppVersion & " - export started"
        WriteLog "document:" & vbTab & MainDocument.FullName
        WriteLog "target folder:" & vbTab & SavePath
        WriteLog "file types:" & vbTab & Join(Filetypes, "; ")
        WriteLog "file name pattern:" & vbTab & Pattern
        WriteLog "if a file already exists:" & vbTab & ExistingMode
        WriteLog "form letters to create:" & vbTab & PlannedCount & " of " & TotalCount
        If StopRecord > 0 Then WriteLog "export stops at record " & StopRecord & " because {" & FieldNames(AbortIndex) & "} is empty"
    End If

        ' do not redraw word during the export; word itself stays visible, otherwise the dialog with
        ' the progress bar and the cancel button would be hidden too
    If IsMac() And Not IsNewMac() Then Me.Hide
    Application.ScreenUpdating = False


    '
    ' save form letter to single files
    '

    UpdateProgress 0, PlannedCount, Tx("PR_STARTING"), EtaLine(PlannedCount)
    Counter = 0

    For i = 1 To CachedRecordCount
        If CancelRequested Then Exit For
        CurrentRecord = RecordNumbers(i)
        TargetName = ""

        If EmptyStopIndex > 0 And i >= EmptyStopIndex Then
            ' stop at the first record where all fields are empty
            StoppedEmpty = True
            WriteLog "record " & CurrentRecord & vbTab & vbTab & "export stopped: all fields are empty"
            Exit For
        End If

        If CurrentRecord = StopRecord Then
            ' stop at the first record where the required field is empty
            Stopped = True
            WriteLog "record " & CurrentRecord & vbTab & vbTab & "export stopped: {" & FieldNames(AbortIndex) & "} is empty"
            Exit For
        End If

        If IsRecordEmpty(i) Then
            ' skip actual record if all columns of the record are empty
            EmptyCount = EmptyCount + 1
            WriteLog "record " & CurrentRecord & vbTab & vbTab & "skipped (empty record)"
        Else
            Counter = Counter + 1
            BaseName = BuildFileBaseName(i, Counter, Pattern, Documentname)
            UpdateProgress Counter - 1, PlannedCount, TxF("PR_LETTER", Counter, PlannedCount, BaseName), EtaLine(PlannedCount - Counter + 1)

            ' determine target files and check for existing files
            ExistingNames = ""
            For t = 0 To TypeCount - 1
                Targets(t) = BaseName & "." & Extensions(t)
                Actions(t) = "saved"
                If FileExists(DirCheckPath & Targets(t)) Then
                    Actions(t) = "exists"
                    ExistingNames = JoinNote(ExistingNames, Targets(t))
                End If
            Next t

            ' handle already existing files
            If ExistingNames <> "" Then
                Mode = ExistingMode
                If Mode = "ask" Then Mode = AskForExistingFile(ExistingNames, ExistingMode)

                Select Case Mode
                Case "overwrite"
                    For t = 0 To TypeCount - 1
                        If Actions(t) = "exists" Then Actions(t) = "overwritten"
                    Next t
                Case "skip"
                    For t = 0 To TypeCount - 1
                        If Actions(t) = "exists" Then Actions(t) = "skipped (file already exists)"
                    Next t
                Case "number"
                    BaseName = GetNumberedBaseName(DirCheckPath, BaseName, Extensions)
                    For t = 0 To TypeCount - 1
                        Targets(t) = BaseName & "." & Extensions(t)
                        Actions(t) = "saved with number"
                    Next t
                Case Else
                    CancelRequested = True
                    WriteLog "record " & CurrentRecord & vbTab & ExistingNames & vbTab & "aborted (file already exists)"
                    Exit For
                End Select
            End If

            AnyToSave = False
            For t = 0 To TypeCount - 1
                If Actions(t) <> "skipped (file already exists)" Then AnyToSave = True
            Next t

            ' create the single document (the time of the form letter is measured for the remaining time)
            LetterStart = Timer
            If AnyToSave Then
                With MainDocument.MailMerge
                    .Destination = wdSendToNewDocument
                    .SuppressBlankLines = True
                    .DataSource.FirstRecord = CurrentRecord
                    .DataSource.LastRecord = CurrentRecord
                    DocumentCount = Documents.Count

                    ' surpress any mailmerge error
                    .Execute Pause:=False
                End With
                If Documents.Count <= DocumentCount Then
                    Err.Raise vbObjectError + 513, , Tx("X_NO_DOC")
                End If
                Set OutputDocument = ActiveDocument

                ' update dynamic fields in the created document
                OutputDocument.Fields.Update
            End If

            ' save the single document in all chosen file types
            For t = 0 To TypeCount - 1
                TargetName = Targets(t)
                If Actions(t) = "skipped (file already exists)" Then
                    SkippedCount = SkippedCount + 1
                Else
                    If Actions(t) = "overwritten" Then
                        Kill SavePath & Targets(t)
                        OverwrittenCount = OverwrittenCount + 1
                    End If
                    SaveDocument OutputDocument, SavePath & Targets(t), CStr(Filetypes(t))
                    FileCount = FileCount + 1
                    If Actions(t) = "saved with number" Then NumberedCount = NumberedCount + 1
                End If
                WriteLog "record " & CurrentRecord & vbTab & Targets(t) & vbTab & Actions(t)
            Next t

            If AnyToSave Then
                ' optional: print document
                If PrintFiles Then
                    If Not PrintDialogShown Then
                        ' the print dialog is shown once, cancelling it stops printing
                        PrintDialogShown = True
                        PrintFiles = (Dialogs(wdDialogFilePrint).Show = -1)
                        SkipSample = True
                    Else
                        OutputDocument.PrintOut Background:=False
                    End If
                End If

                ' close the single document
                OutputDocument.Close SaveChanges:=wdDoNotSaveChanges
                Set OutputDocument = Nothing
                LetterCount = LetterCount + 1

                ' the time of this form letter improves the remaining time (not if the user was in the print dialog)
                If Not SkipSample Then EtaAddSample ElapsedSince(LetterStart)
                SkipSample = False
            End If

            UpdateProgress Counter, PlannedCount, TxF("PR_LETTER", Counter, PlannedCount, BaseName), EtaLine(PlannedCount - Counter)
        End If
    Next i


    '
    ' clean up, unhide word application and show result and error messages
    '

Finish:
    On Error Resume Next

    If Not OutputDocument Is Nothing Then OutputDocument.Close SaveChanges:=wdDoNotSaveChanges
    If Not MainDocument Is Nothing Then
        With MainDocument.MailMerge
            .DataSource.FirstRecord = SavedFirstRecord
            .DataSource.LastRecord = SavedLastRecord
            .Destination = SavedDestination
            .DataSource.ActiveRecord = SavedActiveRecord
        End With
    End If

    Application.ScreenUpdating = True
    Application.Visible = True

    ResultText = TxF("R_LETTERS", TxC(LetterCount, "N_LETTER"), TxC(FileCount, "N_FILE"))
    If OverwrittenCount > 0 Then ResultText = ResultText & vbNewLine & TxF("R_OVERWRITTEN", TxC(OverwrittenCount, "N_EXISTFILE"))
    If NumberedCount > 0 Then ResultText = ResultText & vbNewLine & TxF("R_NUMBERED", TxC(NumberedCount, "N_FILE"))
    If SkippedCount > 0 Then ResultText = ResultText & vbNewLine & TxF("R_SKIPPED", TxC(SkippedCount, "N_FILE"))
    If EmptyCount > 0 Then ResultText = ResultText & vbNewLine & TxF("R_EMPTY", TxC(EmptyCount, "N_EMPTYREC"))
    ResultText = ResultText & vbNewLine & TxF("R_DURATION", FormatDuration(ElapsedSince(ExportStart)))

    If ErrNumber <> 0 Then WriteLog "error " & ErrNumber & ": " & ErrDescription & " (record " & CurrentRecord & ", " & TargetName & ")"
    If CancelRequested Then WriteLog "export cancelled"
    WriteLog "export finished: " & Replace(ResultText, vbNewLine, " ")
    LogWritten = (LogFileNo <> 0)
    CloseLog
    If LogWritten Then ResultText = ResultText & vbNewLine & vbNewLine & TxF("R_LOG", SavePath & LogFileName)

    IsRunning = False
    SetSaveButton False
    On Error GoTo 0

    If ErrNumber <> 0 Then
        MsgBox GetErrorMessage(ErrNumber, ErrDescription, CurrentRecord, TargetName) & vbNewLine & vbNewLine & ResultText, vbOKOnly + vbCritical
    Else
        If CancelRequested Then
            MsgText = Tx("M_CANCELLED") & vbNewLine & vbNewLine & ResultText
        ElseIf StoppedEmpty Then
            MsgText = TxF("M_STOPPED_EMPTY", RecordRef(RecordNumbers(EmptyStopIndex))) & vbNewLine & vbNewLine & ResultText
        ElseIf Stopped Then
            MsgText = TxF("M_STOPPED_FIELD", RecordRef(StopRecord), FieldNames(AbortIndex)) & vbNewLine
            If ScanStoppedAt = StopRecord Then
                MsgText = MsgText & Tx("M_STOPPED_FIELD_A") & vbNewLine & vbNewLine & ResultText
            Else
                MsgText = MsgText & TxF("M_STOPPED_FIELD_B", LetterCount, TxC(TotalCount, "N_LETTER")) & vbNewLine & vbNewLine & ResultText
            End If
        Else
            MsgText = Tx("M_OK") & vbNewLine & vbNewLine & ResultText
        End If
        ' the result message depends on the setting: always, only if there is something to warn about, or never
        HasWarnings = (CancelRequested Or Stopped Or SkippedCount > 0 Or FileCount = 0)
        ResultMode = GetResultMode()
        ShowResult = (ResultMode = "always" Or (ResultMode = "warnings" And HasWarnings))

        If IsMac() Or FileCount = 0 Then
            If ShowResult Then MsgBox MsgText, vbOKOnly + IIf(CancelRequested Or Stopped, vbExclamation, vbInformation)
        Else
            ' open target folder according to the option
            Select Case GetOpenFolderMode()
            Case "open"
                Shell "explorer.exe """ & SavePath & """", vbNormalFocus
                If ShowResult Then MsgBox MsgText, vbOKOnly + IIf(CancelRequested Or Stopped, vbExclamation, vbInformation)
            Case "no"
                If ShowResult Then MsgBox MsgText, vbOKOnly + IIf(CancelRequested Or Stopped, vbExclamation, vbInformation)
            Case Else
                QuestionText = Tx("Q_OPEN")
                If ShowResult Then QuestionText = MsgText & vbNewLine & vbNewLine & QuestionText
                If MsgBox(QuestionText, vbYesNo + IIf(CancelRequested Or Stopped, vbExclamation, vbQuestion) + vbDefaultButton1) = vbYes Then
                    Shell "explorer.exe """ & SavePath & """", vbNormalFocus
                End If
            End Select
        End If
    End If

    Unload Me
    Exit Sub


ErrorHandling:
    ErrNumber = Err.Number
    ErrDescription = Err.Description
    Resume Finish


FolderError:
    If Err.Number = 75 Then
        MsgBox Tx("X_FOLDER_RIGHTS"), vbOKOnly + vbCritical
    ElseIf Err.Number = 76 Then
        MsgBox Tx("X_FOLDER_INVALID"), vbOKOnly + vbCritical
    Else
        MsgBox TxF("X_FOLDER_CREATE", Err.Number, Err.Description), vbOKOnly + vbCritical
    End If

End Sub


Private Sub label_donate_Click()
    ActiveDocument.FollowHyperlink Address:="http://www.formletter2pdf.com/en/?donate=1", NewWindow:=True
End Sub

Private Sub label_subheadline_Click()
    ActiveDocument.FollowHyperlink Address:="https://github.com/farOverNinethousand/formletter2pdf", NewWindow:=True
End Sub

Private Sub lblOldSite_Click()
    ActiveDocument.FollowHyperlink Address:="http://www.formletter2pdf.com/en/", NewWindow:=True
End Sub


Private Sub PlaceOldSiteLink()
    ' link to the old website: right of the "Updates & documentation" text, or below it if there is no room
    Dim Head As Object
    Dim LeftLimit As Single

    If lblOldSite Is Nothing Then Exit Sub
    On Error Resume Next
    Set Head = Me.Controls("label_subheadline")
    If Head Is Nothing Then Exit Sub
    Head.WordWrap = False
    Head.AutoSize = True

    lblOldSite.Caption = Tx("L_OLDSITE")
    lblOldSite.WordWrap = False
    lblOldSite.AutoSize = True
    lblOldSite.Font.Name = Head.Font.Name
    lblOldSite.Font.Size = Head.Font.Size
    lblOldSite.Font.Underline = True
    lblOldSite.ForeColor = Head.ForeColor
    lblOldSite.BackStyle = Head.BackStyle
    lblOldSite.MousePointer = Head.MousePointer

    ' the left column ends at the right edge of the export button; the link text starts at the left edge of the logo
    LeftLimit = Me.Controls("create").Left + Me.Controls("create").Width
    Head.Left = Me.Controls("label_headline").Left
    If Head.Left + Head.Width + 10 + lblOldSite.Width <= LeftLimit Then
        lblOldSite.Left = Head.Left + Head.Width + 10
        lblOldSite.Top = Head.Top
    Else
        lblOldSite.Left = Head.Left
        lblOldSite.Top = Head.Top + Head.Height
    End If
    lblOldSite.ZOrder 0
    On Error GoTo 0
End Sub


' **********************************************************************************************
' data source
' **********************************************************************************************

Private Sub ReadFieldNames()
    Dim Column As MailMergeDataField
    Dim k As Long

    FieldCount = ActiveDocument.MailMerge.DataSource.DataFields.Count
    If FieldCount > 0 Then
        ReDim FieldNames(1 To FieldCount)
    Else
        ReDim FieldNames(0 To 0)
    End If
    k = 0
    For Each Column In ActiveDocument.MailMerge.DataSource.DataFields
        k = k + 1
        FieldNames(k) = Column.Name
    Next Column
End Sub


Private Sub LoadRecords(ByVal Quick As Boolean)
    ' reads all included records once, so preview and export do not need to move through the data source again
    ' every jump to another record makes word render the form letter again, which is slow with a visible preview
    ' of the results; so the field code view is switched on while reading and restored afterwards
    Dim DataSource As MailMergeDataSource
    Dim Column As MailMergeDataField
    Dim SavedRecord As Long
    Dim SavedFieldCodes As Boolean
    Dim LastRecord As Long
    Dim PreviousRecord As Long
    Dim k As Long

    IsLoading = True
    CancelRequested = False
    cboLanguage.Enabled = False
    ScanStart = Timer
    UpdatePreviewButton
    lblProgress.Caption = Tx("PR_READING")
    lblEta.Caption = ""
    lblProgressBar.Visible = False
    lblCount.ForeColor = vbButtonText
    lblCount.Caption = Tx("S_READING")
    Me.Repaint

    On Error GoTo LoadError
    Application.ScreenUpdating = False
    Set DataSource = ActiveDocument.MailMerge.DataSource

    SavedFieldCodes = ActiveDocument.MailMerge.ViewMailMergeFieldCodes
    ActiveDocument.MailMerge.ViewMailMergeFieldCodes = True

    SavedRecord = DataSource.ActiveRecord
    DataSource.ActiveRecord = wdLastRecord
    LastRecord = DataSource.ActiveRecord
    CachedRecordCount = 0
    RecordsInSource = LastRecord
    ScanStoppedAt = 0
    CachedWithStop = (chkStopEmptyRecord.Value = True)
    CachedAbortIndex = GetAbortFieldIndex()
    CachedQuick = Quick
    QuickLimited = False

    If LastRecord > 0 Then
        ReDim RecordNumbers(1 To LastRecord)
        ReDim RecordValues(1 To LastRecord, 0 To FieldCount)

        ' wdNextRecord skips records which are excluded in the recipient list
        DataSource.ActiveRecord = wdFirstRecord
        Do
            CachedRecordCount = CachedRecordCount + 1
            RecordNumbers(CachedRecordCount) = DataSource.ActiveRecord
            k = 0
            For Each Column In DataSource.DataFields
                k = k + 1
                RecordValues(CachedRecordCount, k) = "" & Column.Value
            Next Column

            ' optional: stop reading at the first record where all fields are empty
            If CachedWithStop And IsRecordEmpty(CachedRecordCount) Then
                ScanStoppedAt = RecordNumbers(CachedRecordCount)
                Exit Do
            End If

            ' optional: stop reading at the first record where the chosen field is empty
            If CachedAbortIndex > 0 And Not IsRecordEmpty(CachedRecordCount) Then
                If Trim(RecordValues(CachedRecordCount, CachedAbortIndex)) = "" Then
                    ScanStoppedAt = RecordNumbers(CachedRecordCount)
                    Exit Do
                End If
            End If

            ' optional: quick scan, read only the first records
            If Quick And CachedRecordCount >= QuickScanLimit Then
                If DataSource.ActiveRecord < LastRecord Then QuickLimited = True
                Exit Do
            End If

            ' show the progress and handle the cancel button
            If CachedRecordCount Mod 5 = 0 Or CachedRecordCount = 1 Then
                lblProgress.Caption = TxF("PR_READING_REC", RecordNumbers(CachedRecordCount), LastRecord)
                lblEta.Caption = ScanEtaLine(RecordNumbers(CachedRecordCount), LastRecord)
                lblProgressBar.Width = ProgressBarWidth * RecordNumbers(CachedRecordCount) / LastRecord
                lblProgressBar.Visible = True
                Me.Repaint
                DoEvents
                If CancelRequested Then Exit Do
            End If

            If DataSource.ActiveRecord >= LastRecord Or CachedRecordCount >= LastRecord Then Exit Do
            PreviousRecord = DataSource.ActiveRecord
            DataSource.ActiveRecord = wdNextRecord
            If DataSource.ActiveRecord = PreviousRecord Then Exit Do
        Loop
    End If

LoadFinish:
    On Error Resume Next
    ActiveDocument.MailMerge.ViewMailMergeFieldCodes = SavedFieldCodes
    DataSource.ActiveRecord = SavedRecord
    Application.ScreenUpdating = True
    lblProgress.Caption = Tx("PR_NOT_STARTED")
    lblEta.Caption = ""
    lblProgressBar.Visible = False
    DataLoaded = Not CancelRequested
    IsLoading = False
    cboLanguage.Enabled = True
    UpdatePreviewButton
    On Error GoTo 0
    Exit Sub

LoadError:
    MsgBox TxF("X_READ", Err.Number, Err.Description), vbOKOnly + vbCritical
    CancelRequested = True
    Resume LoadFinish
End Sub


Private Function IsRecordEmpty(ByVal RecordIndex As Long) As Boolean
    Dim k As Long

    IsRecordEmpty = (FieldCount > 0)
    For k = 1 To FieldCount
        If Trim(RecordValues(RecordIndex, k)) <> "" Then
            IsRecordEmpty = False
            Exit Function
        End If
    Next k
End Function


Private Function GetEmptyStopIndex() As Long
    ' returns the index of the first record where all fields are empty, if the export should stop there
    Dim i As Long

    If chkStopEmptyRecord.Value <> True Then Exit Function
    For i = 1 To CachedRecordCount
        If IsRecordEmpty(i) Then
            GetEmptyStopIndex = i
            Exit Function
        End If
    Next i
End Function


Private Function GetAbortFieldIndex() As Long
    If chkAbortEmpty.Value = True And cboAbortField.ListIndex >= 0 Then
        GetAbortFieldIndex = cboAbortField.ListIndex + 1
    Else
        GetAbortFieldIndex = 0
    End If
End Function


' **********************************************************************************************
' file names and paths
' **********************************************************************************************

Private Function BuildFileBaseName(ByVal RecordIndex As Long, ByVal Counter As Long, ByVal Pattern As String, ByVal Documentname As String) As String
    Dim File As String
    Dim k As Long

    ' replace datasource field markers
    File = Pattern
    For k = 1 To FieldCount
        File = Replace(File, "{" & FieldNames(k) & "}", RecordValues(RecordIndex, k))
    Next k

    ' replace generic counter marker
    File = Replace(File, "{%counter%}", CStr(Counter))

    ' replace document name marker
    File = Replace(File, "{%filename%}", Documentname)

    File = SanitizeFileName(File)
    If File = "" Then File = "formletter-" & Counter
    BuildFileBaseName = File
End Function


Private Function SanitizeFileName(ByVal Text As String) As String
    Dim Character As Variant
    Dim i As Long

    ' remove all characters which are not allowed in file names (Chr(34) = ")
    For Each Character In Array("\", "/", ":", "*", "?", Chr(34), "<", ">", "|")
        Text = Replace(Text, Character, "")
    Next Character
    For i = 0 To 31
        Text = Replace(Text, Chr(i), "")
    Next i

    Text = Trim(Text)
    If Len(Text) > MaxFileNameLength Then Text = Left(Text, MaxFileNameLength)

    ' file names must not end with a dot or a space on windows
    Do While Len(Text) > 0 And (Right(Text, 1) = "." Or Right(Text, 1) = " ")
        Text = Left(Text, Len(Text) - 1)
    Loop

    SanitizeFileName = Text
End Function


Private Function FindUnknownMarker(ByVal Pattern As String) As String
    ' returns the first marker of the pattern which is neither a data source field nor a generic marker
    Dim StartPos As Long
    Dim EndPos As Long
    Dim k As Long

    For k = 1 To FieldCount
        Pattern = Replace(Pattern, "{" & FieldNames(k) & "}", "")
    Next k
    Pattern = Replace(Pattern, "{%counter%}", "")
    Pattern = Replace(Pattern, "{%filename%}", "")

    StartPos = InStr(Pattern, "{")
    If StartPos > 0 Then
        EndPos = InStr(StartPos, Pattern, "}")
        If EndPos > StartPos Then FindUnknownMarker = Mid(Pattern, StartPos, EndPos - StartPos + 1)
    End If
End Function


Private Function HasInvalidCharacters(ByVal Text As String) As Boolean
    Dim Character As Variant

    For Each Character In Array("\", "/", ":", "*", "?", Chr(34), "<", ">", "|")
        If InStr(Text, Character) > 0 Then
            HasInvalidCharacters = True
            Exit Function
        End If
    Next Character
End Function


Private Function GetDocumentName(Doc As Document) As String
    If InStrRev(Doc.Name, ".") <> 0 Then
        GetDocumentName = Left(Doc.Name, InStrRev(Doc.Name, ".") - 1)
    Else
        GetDocumentName = Doc.Name
    End If
End Function


Private Function CleanLocation(ByVal Location As String) As String
    ' remove trailing backslash on windows, e.g. "D:\"
    If Not IsMac() And Len(Location) > 1 And Right(Location, 1) = "\" Then Location = Left(Location, Len(Location) - 1)
    CleanLocation = Location
End Function


Private Sub GetTargetPaths(ByVal Location As String, ByVal Folder As String, ByRef SavePath As String, ByRef DirCheckPath As String)
    ' SavePath is used to save files, DirCheckPath to check for existing files and folders
    If IsNewMac() Then
        SavePath = "/Volumes/" & Replace(Location, ":", "/") & Folder & "/"
        DirCheckPath = Location & Folder & ":"
    ElseIf IsMac() Then
        SavePath = Location & Folder & ":"
        DirCheckPath = SavePath
    Else
        SavePath = Location & "\" & Folder & "\"
        DirCheckPath = SavePath
    End If
End Sub


Private Function TargetFolderExists(ByVal SavePath As String, ByVal DirCheckPath As String) As Boolean
    If IsNewMac() Then
        TargetFolderExists = (Len(Dir(DirCheckPath, vbDirectory)) > 0)
    ElseIf IsMac() Then
        TargetFolderExists = (MacScript("tell application " & Chr(34) & "Finder" & Chr(34) & Chr(13) & "exists folder " & Chr(34) & SavePath & Chr(34) & Chr(13) & "end tell" & Chr(13)) = True)
    Else
        TargetFolderExists = (Len(Dir(SavePath, vbDirectory)) > 0)
    End If
End Function


Private Sub CreateTargetFolder(ByVal Location As String, ByVal Folder As String, ByVal SavePath As String)
    Dim fso As Object

    If IsNewMac() Then
        MkDir SavePath
    ElseIf IsMac() Then
        MacScript ("tell application " & Chr(34) & "Finder" & Chr(34) & Chr(13) & "make new folder at " & Chr(34) & Location & Chr(34) & " with properties {name: " & Chr(34) & Folder & Chr(34) & "}" & Chr(13) & "end tell" & Chr(13))
    Else
        Set fso = CreateObject("Scripting.FileSystemObject")
        fso.CreateFolder SavePath
    End If
End Sub


Private Function FileExists(ByVal FullPath As String) As Boolean
    On Error Resume Next
    FileExists = (Len(Dir(FullPath, vbNormal + vbReadOnly + vbHidden + vbSystem)) > 0)
    On Error GoTo 0
End Function


Private Function GetNumberedBaseName(ByVal DirCheckPath As String, ByVal BaseName As String, Extensions() As String) As String
    ' finds the first "name (n)" for which no file of any chosen file type exists
    Dim Number As Long
    Dim Candidate As String
    Dim IsFree As Boolean
    Dim t As Long

    Number = 2
    Do
        Candidate = BaseName & " (" & Number & ")"
        IsFree = True
        For t = LBound(Extensions) To UBound(Extensions)
            If FileExists(DirCheckPath & Candidate & "." & Extensions(t)) Then
                IsFree = False
                Exit For
            End If
        Next t
        If IsFree Then Exit Do
        Number = Number + 1
    Loop
    GetNumberedBaseName = Candidate
End Function


' **********************************************************************************************
' file types
' **********************************************************************************************

Private Sub BuildFiletypeList()
    ' replaces the option buttons and the select box of the form layout with a list of check boxes
    Dim CtlName As Variant
    Dim Ctl As MSForms.Control
    Dim BoxLeft As Single
    Dim BoxTop As Single
    Dim BoxRight As Single
    Dim BoxBottom As Single
    Dim Needed As Single
    Dim NextTop As Single
    Dim Delta As Single

    BoxLeft = 100000
    BoxTop = 100000
    For Each CtlName In Array("filetype_pdf", "filetype_docx", "filetype_other", "filetype_other_select")
        Set Ctl = Me.Controls(CStr(CtlName))
        If Ctl.Left < BoxLeft Then BoxLeft = Ctl.Left
        If Ctl.Top < BoxTop Then BoxTop = Ctl.Top
        If Ctl.Left + Ctl.Width > BoxRight Then BoxRight = Ctl.Left + Ctl.Width
        If Ctl.Top + Ctl.Height > BoxBottom Then BoxBottom = Ctl.Top + Ctl.Height
        Ctl.Visible = False
    Next CtlName

    ' one check box per file type instead of a list box: a list box shows a scroll bar when its
    ' height does not fit the row height of the font, check boxes always show all entries
    FiletypeLeft = BoxLeft
    FiletypeTop = BoxTop
    FiletypeWidth = BoxRight - BoxLeft
    FiletypeCount = 0

    ' order of the entries is the order the files are saved: fixed formats first, lossy formats last
    AddFiletype "Portable Document Format (.pdf)", "pdf"
    If Not IsMac() Then
        ' pdf export settings needed for this document type not available in office for mac
        AddFiletype "PDF/A (ISO 19005-1, .pdf)", "PDF/A (ISO 19005-1-compatible, .pdf)"
    End If
    AddFiletype "Word Document (.docx)", "docx"
    AddFiletype "Word 97-2003 Document (.doc)", "Word 97-2003 Document (.doc)"
    AddFiletype "Rich Text Format (.rtf)", "Rich Text Format (.rtf)"
    AddFiletype "Single File Web Page (.mht)", "Single File Web Page (.mht)"
    AddFiletype "Web Page (.htm)", "Web Page (.htm)"
    AddFiletype "Web Page, Filtered (.htm)", "Web Page, Filtered (.htm)"
    AddFiletype "Plain Text (.txt)", "Plain Text (.txt)"

    ' use the free space below the entries (up to the next visible control);
    ' only if it is not enough, move the controls below down
    Needed = FiletypeCount * FiletypeRowHeight + 4
    NextTop = 100000
    For Each Ctl In Me.Controls
        If Ctl.Visible And Left(Ctl.Name, 6) <> "ext_ft" Then
            If Ctl.Top >= BoxBottom - 1 And Ctl.Left < BoxRight And Ctl.Left + Ctl.Width > BoxLeft Then
                If Ctl.Top < NextTop Then NextTop = Ctl.Top
            End If
        End If
    Next Ctl

    If NextTop < 100000 Then
        Delta = Needed - (NextTop - BoxTop - 4)
    Else
        Delta = 0
    End If

    If Delta > 0 Then
        For Each Ctl In Me.Controls
            If Not TypeOf Ctl.Parent Is MSForms.Frame Then
                If Ctl.Top >= BoxBottom - 1 And Left(Ctl.Name, 6) <> "ext_ft" Then Ctl.Top = Ctl.Top + Delta
            End If
        Next Ctl
        Me.Height = Me.Height + Delta
    End If
End Sub


Private Sub BuildMessagesSection()
    ' adds the section "Messages" with the settings for the confirm dialog, the result message and
    ' (windows only) opening the output folder directly above the create button;
    ' the create button and the controls below move down
    Dim Ctl As MSForms.Control
    Dim Fra As MSForms.Frame
    Dim PosTop As Single
    Dim FrameHeight As Single
    Dim FrameLeft As Single
    Dim FrameWidth As Single

    FrameHeight = MessagesFrameHeight
    If IsMac() Then FrameHeight = FrameHeight - 24

    PosTop = Me.Controls("create").Top
    For Each Ctl In Me.Controls
        If Not TypeOf Ctl.Parent Is MSForms.Frame And Left(Ctl.Name, 4) <> "ext_" Then
            If Ctl.Top >= PosTop - 1 Then Ctl.Top = Ctl.Top + FrameHeight + 6
        End If
    Next Ctl
    Me.Height = Me.Height + FrameHeight + 6

    FrameLeft = Me.Controls("label_foldername").Left
    FrameWidth = Me.Controls("foldername").Left + Me.Controls("foldername").Width - FrameLeft
    Set Fra = AddControl(Me, "Forms.Frame.1", "ext_frame_messages", FrameLeft, PosTop, FrameWidth, FrameHeight)
    Set fraMessages = Fra

    ' the texts of the labels and the entries of the lists are set in ApplyLanguage
    Set cboConfirm = AddMessageCombo(Fra, "ext_confirm", lblConfirm, 4, FrameWidth)
    FillCombo cboConfirm, "M_ALWAYS,M_WARNINGS,M_NEVER"
    SelectMode cboConfirm, MessageModes, IIf(StorageConfirmBeforeWriting <> "", StorageConfirmBeforeWriting, DefaultConfirmBeforeWriting)

    Set cboResult = AddMessageCombo(Fra, "ext_result", lblResultMsg, 28, FrameWidth)
    FillCombo cboResult, "M_ALWAYS,M_WARNINGS,M_NEVER"
    SelectMode cboResult, MessageModes, IIf(StorageResultMessage <> "", StorageResultMessage, DefaultResultMessage)

    If Not IsMac() Then
        Set cboOpenFolder = AddMessageCombo(Fra, "ext_open_folder", lblOpenFolder, 52, FrameWidth)
        FillCombo cboOpenFolder, "M_ASK,M_OPEN,M_NOOPEN"
        SelectMode cboOpenFolder, OpenFolderModes, IIf(StorageOpenFolder <> "", StorageOpenFolder, DefaultOpenFolder)
    End If
End Sub


Private Function AddMessageCombo(Fra As MSForms.Frame, ByVal CtlName As String, ByRef Lbl As MSForms.Label, ByVal PosTop As Single, ByVal FrameWidth As Single) As MSForms.ComboBox
    ' label on the left and a drop-down list on the right of the section
    Dim Box As MSForms.ComboBox

    Set Lbl = AddControl(Fra, "Forms.Label.1", CtlName & "_label", 6, PosTop + 3, 160, 12)
    Set Box = AddControl(Fra, "Forms.ComboBox.1", CtlName, 172, PosTop, FrameWidth - 172 - 12, 18)
    Box.Style = fmStyleDropDownList
    Set AddMessageCombo = Box
End Function

Private Sub SelectMode(Box As MSForms.ComboBox, ByVal Modes As String, ByVal Value As String)
    ' selects the entry whose internal name (list "Modes", same order as the entries) is Value; first entry otherwise
    Dim ModeList As Variant
    Dim i As Long

    Box.ListIndex = 0
    ModeList = Split(Modes, ",")
    For i = 0 To UBound(ModeList)
        If LCase(Value) = ModeList(i) Then Box.ListIndex = i
    Next i
End Sub


Private Function GetMessageMode(Box As MSForms.ComboBox, ByVal Modes As String) As String
    Dim ModeList As Variant

    ModeList = Split(Modes, ",")
    If Box Is Nothing Then
        GetMessageMode = ModeList(0)
    ElseIf Box.ListIndex < 0 Then
        GetMessageMode = ModeList(0)
    Else
        GetMessageMode = ModeList(Box.ListIndex)
    End If
End Function


Private Function GetConfirmMode() As String
    GetConfirmMode = GetMessageMode(cboConfirm, MessageModes)
End Function


Private Function GetResultMode() As String
    GetResultMode = GetMessageMode(cboResult, MessageModes)
End Function

Private Function GetOpenFolderMode() As String
    If cboOpenFolder Is Nothing Then
        GetOpenFolderMode = "ask"
    ElseIf cboOpenFolder.ListIndex < 0 Then
        GetOpenFolderMode = "ask"
    Else
        GetOpenFolderMode = Split(OpenFolderModes, ",")(cboOpenFolder.ListIndex)
    End If
End Function


Private Sub AddFiletype(ByVal Caption As String, ByVal Id As String)
    ' creates the check box of the file type; the events are handled by the module level variables chkFt1 to chkFt10
    Dim Box As MSForms.CheckBox

    Set Box = AddControl(Me, "Forms.CheckBox.1", "ext_ft" & (FiletypeCount + 1), FiletypeLeft, FiletypeTop + FiletypeCount * FiletypeRowHeight, FiletypeWidth, FiletypeRowHeight - 1)
    Box.Caption = Caption
    Box.WordWrap = False

    Select Case FiletypeCount
    Case 0: Set chkFt1 = Box
    Case 1: Set chkFt2 = Box
    Case 2: Set chkFt3 = Box
    Case 3: Set chkFt4 = Box
    Case 4: Set chkFt5 = Box
    Case 5: Set chkFt6 = Box
    Case 6: Set chkFt7 = Box
    Case 7: Set chkFt8 = Box
    Case 8: Set chkFt9 = Box
    Case 9: Set chkFt10 = Box
    End Select

    ReDim Preserve FiletypeIds(0 To FiletypeCount)
    FiletypeIds(FiletypeCount) = Id
    FiletypeCount = FiletypeCount + 1
End Sub


Private Function GetFiletypeBox(ByVal Index As Long) As MSForms.CheckBox
    Select Case Index
    Case 0: Set GetFiletypeBox = chkFt1
    Case 1: Set GetFiletypeBox = chkFt2
    Case 2: Set GetFiletypeBox = chkFt3
    Case 3: Set GetFiletypeBox = chkFt4
    Case 4: Set GetFiletypeBox = chkFt5
    Case 5: Set GetFiletypeBox = chkFt6
    Case 6: Set GetFiletypeBox = chkFt7
    Case 7: Set GetFiletypeBox = chkFt8
    Case 8: Set GetFiletypeBox = chkFt9
    Case 9: Set GetFiletypeBox = chkFt10
    End Select
End Function


Private Function GetSelectedFiletypes() As String
    ' returns the chosen file types separated by ";" in the order they are saved, e.g. "pdf;docx"
    Dim i As Long

    For i = 0 To FiletypeCount - 1
        If GetFiletypeBox(i).Value = True Then
            If GetSelectedFiletypes <> "" Then GetSelectedFiletypes = GetSelectedFiletypes & ";"
            GetSelectedFiletypes = GetSelectedFiletypes & FiletypeIds(i)
        End If
    Next i
End Function


Private Function UniqueExtensions(Extensions() As String) As String()
    ' every file extension counts only once; file types with the same extension are rejected with an error when exporting
    Dim Result() As String
    Dim Seen As New Collection
    Dim Count As Long
    Dim t As Long

    ReDim Result(0 To UBound(Extensions))
    For t = LBound(Extensions) To UBound(Extensions)
        If AddToSet(Seen, Extensions(t)) Then
            Result(Count) = Extensions(t)
            Count = Count + 1
        End If
    Next t
    ReDim Preserve Result(0 To Count - 1)
    UniqueExtensions = Result
End Function


Private Function RecordRef(ByVal RecordNumber As Long) As String
    ' "record 17/150": record number and total number of records of the data source
    If RecordsInSource > 0 Then
        RecordRef = TxF("REC_REF", RecordNumber, RecordsInSource)
    Else
        RecordRef = TxF("REC_REF_SHORT", RecordNumber)
    End If
End Function


Private Function GetExtensions(Filetypes As Variant) As String()
    Dim Extensions() As String
    Dim t As Long

    ReDim Extensions(0 To UBound(Filetypes))
    For t = 0 To UBound(Filetypes)
        Extensions(t) = GetExtension(CStr(Filetypes(t)))
    Next t
    GetExtensions = Extensions
End Function


Private Function GetExtension(ByVal Filetype As String) As String
    Select Case Filetype
    Case "docx"
        GetExtension = "docx"
    Case "Word 97-2003 Document (.doc)"
        GetExtension = "doc"
    Case "Single File Web Page (.mht)"
        GetExtension = "mht"
    Case "Web Page (.htm)", "Web Page, Filtered (.htm)"
        GetExtension = "htm"
    Case "Rich Text Format (.rtf)"
        GetExtension = "rtf"
    Case "Plain Text (.txt)"
        GetExtension = "txt"
    Case Else
        GetExtension = "pdf"
    End Select
End Function


Private Sub SaveDocument(Doc As Document, ByVal Targetfile As String, ByVal Filetype As String)
    If Filetype = "docx" Then
        Doc.SaveAs Filename:=Targetfile, FileFormat:=wdFormatDocumentDefault
    ElseIf Filetype = "PDF/A (ISO 19005-1-compatible, .pdf)" Then
        Doc.ExportAsFixedFormat OutputFileName:=Targetfile, ExportFormat:=wdExportFormatPDF, OpenAfterExport:=False, OptimizeFor:=wdExportOptimizeForPrint, Range:=wdExportAllDocument, Item:=wdExportDocumentContent, IncludeDocProps:=True, KeepIRM:=True, CreateBookmarks:=wdExportCreateNoBookmarks, DocStructureTags:=True, BitmapMissingFonts:=True, UseISO19005_1:=True
    ElseIf Filetype = "Word 97-2003 Document (.doc)" Then
        Doc.SaveAs Filename:=Targetfile, FileFormat:=wdFormatDocument97
    ElseIf Filetype = "Single File Web Page (.mht)" Then
        Doc.SaveAs Filename:=Targetfile, FileFormat:=wdFormatWebArchive
    ElseIf Filetype = "Web Page (.htm)" Then
        Doc.SaveAs Filename:=Targetfile, FileFormat:=wdFormatHTML
    ElseIf Filetype = "Web Page, Filtered (.htm)" Then
        Doc.SaveAs Filename:=Targetfile, FileFormat:=wdFormatFilteredHTML
    ElseIf Filetype = "Rich Text Format (.rtf)" Then
        Doc.SaveAs Filename:=Targetfile, FileFormat:=wdFormatRTF
    ElseIf Filetype = "Plain Text (.txt)" Then
        Doc.SaveAs Filename:=Targetfile, FileFormat:=wdFormatEncodedText
    Else
        ' simple command for default print quality - no better quality than this available
        Doc.SaveAs Filename:=Targetfile, FileFormat:=wdFormatPDF

        ' alternativ method:
        ' extended command to export document as pdf, with the possibility to set different export options
        'Doc.ExportAsFixedFormat OutputFileName:=Targetfile, ExportFormat:=wdExportFormatPDF, OpenAfterExport:=False, OptimizeFor:=wdExportOptimizeForPrint, Range:=wdExportAllDocument, Item:=wdExportDocumentContent, IncludeDocProps:=True, KeepIRM:=True, CreateBookmarks:=wdExportCreateWordBookmarks, DocStructureTags:=True, BitmapMissingFonts:=True, UseISO19005_1:=False
    End If
End Sub


' **********************************************************************************************
' preview
' **********************************************************************************************

Private Sub RefreshPreview()
    Dim Filetypes As Variant
    Dim Extensions() As String
    Dim AllExtensions() As String
    Dim Pattern As String
    Dim Documentname As String
    Dim SavePath As String
    Dim DirCheckPath As String
    Dim BaseName As String
    Dim TargetName As String
    Dim Note As String
    Dim ModeText As String
    Dim CountText As String
    Dim Marker As String
    Dim CheckExisting As Boolean
    Dim LongestExt As Long
    Dim Seen As Collection
    Dim TypeCount As Long
    Dim AbortIndex As Long
    Dim StopRecord As Long
    Dim EmptyStopIndex As Long
    Dim i As Long
    Dim t As Long
    Dim Counter As Long
    Dim TotalCount As Long
    Dim EmptyCount As Long
    Dim ConflictCount As Long
    Dim NoteDup As Boolean
    Dim NoteExists As Boolean

    If IsInitializing Or IsRunning Or IsLoading Or Not DataLoaded Or lstPreview Is Nothing Then Exit Sub

    lstPreview.Clear
    ShowPreviewDetail
    lblCount.ForeColor = vbRed
    If GetSelectedFiletypes() = "" Then
        lblCount.Caption = Tx("E_NO_FILETYPE")
        SetCreateCaption -1
        Exit Sub
    End If

    Filetypes = Split(GetSelectedFiletypes(), ";")
    AllExtensions = GetExtensions(Filetypes)
    Extensions = UniqueExtensions(AllExtensions)
    TypeCount = UBound(Extensions) + 1
    Pattern = Me.filenamepattern.Value
    Documentname = GetDocumentName(ActiveDocument)
    ModeText = Tx("MODE_" & UCase(GetExistingMode()))
    AbortIndex = GetAbortFieldIndex()

    ' existing files can only be detected if the target folder is known
    If Me.filelocation.Value <> "" And Trim(Me.foldername.Value) <> "" Then
        CheckExisting = True
        GetTargetPaths CleanLocation(Me.filelocation.Value), Replace(Me.foldername.Value, "{%timestamp%}", Format(Now, TimestampFormat)), SavePath, DirCheckPath
    End If

    Set Seen = New Collection
    EmptyStopIndex = GetEmptyStopIndex()
    LongestExt = 0
    For t = 0 To TypeCount - 1
        If Len(Extensions(t)) > LongestExt Then LongestExt = Len(Extensions(t))
    Next t

    For i = 1 To CachedRecordCount
        If EmptyStopIndex > 0 And i >= EmptyStopIndex Then
            AddPreviewRow RecordNumbers(i), "", TxF("P_ALL_EMPTY_STOP", RecordRef(RecordNumbers(i)))
            Exit For
        End If
        If IsRecordEmpty(i) Then
            EmptyCount = EmptyCount + 1
            AddPreviewRow RecordNumbers(i), "", Tx("P_EMPTY_SKIPPED")
        Else
            TotalCount = TotalCount + 1
            If StopRecord = 0 And AbortIndex > 0 Then
                If Trim(RecordValues(i, AbortIndex)) = "" Then
                    StopRecord = RecordNumbers(i)
                    AddPreviewRow RecordNumbers(i), "", TxF("P_FIELD_STOP", FieldNames(AbortIndex), RecordRef(RecordNumbers(i)))
                End If
            End If

            If StopRecord > 0 Then
                If StopRecord <> RecordNumbers(i) Then AddPreviewRow RecordNumbers(i), "", Tx("P_NOT_EXPORTED")
            Else
                Counter = Counter + 1
                BaseName = BuildFileBaseName(i, Counter, Pattern, Documentname)
                Note = ""
                NoteDup = False
                NoteExists = False
                If IsReservedName(BaseName) Then Note = JoinNote(Note, Tx("P_RESERVED"))
                If CheckExisting And Not IsMac() Then
                    If Len(SavePath) + Len(BaseName) + 1 + LongestExt > MaxPathLength Then Note = JoinNote(Note, Tx("P_PATH_LONG"))
                End If
                For t = 0 To TypeCount - 1
                    TargetName = BaseName & "." & Extensions(t)
                    If Not AddToSet(Seen, TargetName) Then
                        ConflictCount = ConflictCount + 1
                        If Not NoteDup Then
                            Note = JoinNote(Note, TxF("P_DUP", ModeText))
                            NoteDup = True
                        End If
                    ElseIf CheckExisting Then
                        If FileExists(DirCheckPath & TargetName) Then
                            ConflictCount = ConflictCount + 1
                            If Not NoteExists Then
                                Note = JoinNote(Note, TxF("P_EXISTS", ModeText))
                                NoteExists = True
                            End If
                        End If
                    End If
                Next t
                AddPreviewRow RecordNumbers(i), BaseName & "." & Join(Extensions, " + ."), Note
            End If
        End If
    Next i

    PreviewConflicts = ConflictCount

    ' the create button shows how many files are written; not known after a quick scan
    If QuickLimited Then
        SetCreateCaption -1
    ElseIf GetExistingMode() = "skip" Then
        SetCreateCaption Counter * TypeCount - ConflictCount
    Else
        SetCreateCaption Counter * TypeCount
    End If

    ' quick scan: tell the user that further records exist which are not shown
    If QuickLimited Then
        lstPreview.AddItem "..."
        lstPreview.List(lstPreview.ListCount - 1, 1) = TxF("P_QUICK_ROW", CachedRecordCount, RecordsInSource)
        lstPreview.List(lstPreview.ListCount - 1, 2) = ""
    End If

    ' summary above the list
    lblCount.ForeColor = vbButtonText
    If GetExistingMode() = "skip" Then
        CountText = TxF("P_SUM_CREATED", TxC(Counter, "N_LETTER"), TxC(Counter * TypeCount - ConflictCount, "N_FILE"))
    Else
        CountText = TxF("P_SUM_CREATED", TxC(Counter, "N_LETTER"), TxC(Counter * TypeCount, "N_FILE"))
    End If
    If QuickLimited Then CountText = TxF("P_QUICK_SUM", CachedRecordCount, RecordsInSource) & CountText
    If StopRecord > 0 Then
        lblCount.ForeColor = vbRed
        If ScanStoppedAt = StopRecord Then
            CountText = CountText & TxF("P_STOPS_FIELD", RecordRef(StopRecord), FieldNames(AbortIndex))
        Else
            CountText = CountText & TxF("P_STOPS_FIELD_FOUND", RecordRef(StopRecord), FieldNames(AbortIndex), TxC(TotalCount, "N_LETTER"))
        End If
    End If
    If EmptyStopIndex > 0 Then
        lblCount.ForeColor = vbRed
        CountText = CountText & TxF("P_STOPS_EMPTY", RecordRef(RecordNumbers(EmptyStopIndex)))
    End If
    If ConflictCount > 0 Then CountText = CountText & TxC(ConflictCount, "P_CONFLICT", ModeText)
    If EmptyCount > 0 Then CountText = CountText & TxC(EmptyCount, "P_EMPTYSKIP")

    Marker = FindUnknownMarker(Pattern)
    If Marker <> "" Then
        lblCount.ForeColor = vbRed
        CountText = CountText & vbNewLine & TxF("P_UNKNOWN_MARKER", Marker)
    End If
    lblCount.Caption = CountText
End Sub


Private Sub AddPreviewRow(ByVal RecordNumber As Long, ByVal FileName As String, ByVal Note As String)
    lstPreview.AddItem CStr(RecordNumber)
    If FileName = "" Then
        ' rows without a file name (skipped records, stop of the scan or export): the note takes the place of the file name
        lstPreview.List(lstPreview.ListCount - 1, 1) = Note
        lstPreview.List(lstPreview.ListCount - 1, 2) = ""
    Else
        lstPreview.List(lstPreview.ListCount - 1, 1) = FileName
        lstPreview.List(lstPreview.ListCount - 1, 2) = Note
    End If
End Sub


Private Function AddToSet(Seen As Collection, ByVal Key As String) As Boolean
    ' returns False if the key already exists (collection keys are case-insensitive, like windows file names)
    On Error Resume Next
    Seen.Add True, Key
    AddToSet = (Err.Number = 0)
    On Error GoTo 0
End Function


Private Function JoinNote(ByVal Note As String, ByVal Addition As String) As String
    If Note = "" Then
        JoinNote = Addition
    Else
        JoinNote = Note & ", " & Addition
    End If
End Function


' **********************************************************************************************
' export state, progress and questions during the export
' **********************************************************************************************

Private Function GetExistingMode() As String
    If cboExisting.ListIndex < 0 Then
        GetExistingMode = "ask"
    Else
        GetExistingMode = Split(ExistingModes, ",")(cboExisting.ListIndex)
    End If
End Function


Private Sub StartRunning()
    Dim Ctl As MSForms.Control

    IsRunning = True
    CancelRequested = False
    btnPreview.Enabled = False
    SetSaveButton True
    EtaSum = 0
    EtaCount = 0
    ExportStart = Timer
    lblEta.Caption = ""
    PriorSecondsPerLetter = DefaultSecondsPerPage * MainDocumentPages()
    If PriorSecondsPerLetter <= 0 Then PriorSecondsPerLetter = 1

    ' only cancel, progress and questions stay usable during the export
    For Each Ctl In Me.Controls
        If Not TypeOf Ctl.Parent Is MSForms.Frame Then
            Select Case Ctl.Name
            Case "cancel", "create", "ext_frame_progress", "ext_frame_conflict"
            Case Else
                Ctl.Enabled = False
            End Select
        End If
    Next Ctl
End Sub


Private Sub RequestCancel()
    CancelRequested = True
    lblProgress.Caption = IIf(IsLoading, Tx("PR_CANCEL_SCAN"), Tx("PR_CANCEL_SAVE"))
End Sub


Private Sub UpdateProgress(ByVal DoneCount As Long, ByVal TotalCount As Long, ByVal Text As String, Optional ByVal EtaText As String = "")
    If CancelRequested Then
        Text = Tx("PR_CANCEL_SAVE")
        EtaText = ""
    End If
    lblProgress.Caption = Text
    lblEta.Caption = EtaText
    If TotalCount > 0 And DoneCount > 0 Then
        lblProgressBar.Width = ProgressBarWidth * DoneCount / TotalCount
        lblProgressBar.Visible = True
    Else
        lblProgressBar.Visible = False
    End If
    Me.Repaint
    DoEvents
End Sub


Private Function ElapsedSince(ByVal StartTimer As Double) As Double
    ' seconds since the Timer value StartTimer; the timer starts again at midnight
    ElapsedSince = Timer - StartTimer
    If ElapsedSince < 0 Then ElapsedSince = ElapsedSince + 86400
End Function


Private Sub EtaAddSample(ByVal Seconds As Double)
    EtaSum = EtaSum + Seconds
    EtaCount = EtaCount + 1
End Sub


Private Function MainDocumentPages() As Long
    ' number of pages of the main document, taken as the number of pages of one form letter (at least 1)
    On Error Resume Next
    MainDocumentPages = ActiveDocument.ComputeStatistics(wdStatisticPages)
    On Error GoTo 0
    If MainDocumentPages < 1 Then MainDocumentPages = 1
End Function


Private Function EtaLine(ByVal RemainingLetters As Long) As String
    ' remaining time of the export: the estimate per form letter starts with DefaultSecondsPerPage * pages and counts
    ' like EtaPriorWeight measured form letters; with every created form letter the measured times count more
    Dim Prior As Double
    Dim Average As Double

    If RemainingLetters <= 0 Then Exit Function
    Prior = PriorSecondsPerLetter
    Average = (Prior * EtaPriorWeight + EtaSum) / (EtaPriorWeight + EtaCount)
    EtaLine = TxF("ETA_REMAINING", FormatDuration(RoundEta(Average * RemainingLetters)))
End Function


Private Function ScanEtaLine(ByVal RecordsDone As Long, ByVal RecordsTotal As Long) As String
    ' remaining time of the scan, only from measured times (after some records and some time)
    Dim Elapsed As Double

    Elapsed = ElapsedSince(ScanStart)
    If RecordsDone < 5 Or Elapsed < 2 Or RecordsTotal <= RecordsDone Then Exit Function
    ScanEtaLine = TxF("ETA_REMAINING", FormatDuration(RoundEta(Elapsed / RecordsDone * (RecordsTotal - RecordsDone))))
End Function


Private Function RoundEta(ByVal Seconds As Double) As Double
    ' under a minute in steps of 5 seconds (at least 5), otherwise in steps of 10 seconds
    If Seconds < 60 Then
        RoundEta = Int((Seconds + 4) / 5) * 5
        If RoundEta < 5 Then RoundEta = 5
    Else
        RoundEta = Int((Seconds + 5) / 10) * 10
    End If
End Function


Private Function FormatDuration(ByVal Seconds As Double) As String
    Dim Total As Long

    Total = CLng(Seconds)
    If Total < 60 Then
        FormatDuration = TxF("DUR_S", Total)
    ElseIf Total < 3600 Then
        FormatDuration = TxF("DUR_MS", Total \ 60, Total Mod 60)
    Else
        FormatDuration = TxF("DUR_HM", Total \ 3600, (Total Mod 3600) \ 60)
    End If
End Function

Private Function AskForExistingFile(ByVal FileNames As String, ByRef ExistingMode As String) As String
    ' asks what to do with existing files; with "apply to all" the answer is used for all further files

    If Not Me.Visible Then
        ' the form is hidden on office for mac 2011
        Select Case MsgBox(TxF("CONFLICT_MAC", FileNames), vbYesNoCancel + vbQuestion)
        Case vbYes
            AskForExistingFile = "overwrite"
        Case vbNo
            AskForExistingFile = "skip"
        Case Else
            AskForExistingFile = "number"
        End Select
        Exit Function
    End If

    lblConflict.Caption = TxF("CONFLICT_TEXT", FileNames)
    chkConflictAll.Value = False
    ConflictChoice = ""
    fraPreview.Visible = False
    fraConflict.Visible = True

    Do While ConflictChoice = "" And Not CancelRequested
        DoEvents
        WaitShortly
    Loop
    If ConflictChoice = "" Then ConflictChoice = "abort"

    If chkConflictAll.Value = True And ConflictChoice <> "abort" Then ExistingMode = ConflictChoice

    fraConflict.Visible = False
    fraPreview.Visible = True
    AskForExistingFile = ConflictChoice
End Function


Private Sub WaitShortly()
#If Mac Then
    ' DoEvents only
#Else
    Sleep 20
#End If
End Sub


Private Function GetErrorMessage(ByVal ErrNumber As Long, ByVal ErrDescription As String, ByVal RecordNumber As Long, ByVal FileName As String) As String
    Dim Context As String

    If FileName <> "" Then Context = vbNewLine & vbNewLine & TxF("X_CONTEXT", RecordNumber, FileName)

    Select Case ErrNumber
    Case 4198
        GetErrorMessage = Tx("X_4198") & Context
    Case 5852
        GetErrorMessage = Tx("X_5852")
    Case 70, 75
        GetErrorMessage = Tx("X_70") & Context
    Case 76
        GetErrorMessage = Tx("X_76")
    Case Else
        GetErrorMessage = TxF("X_UNKNOWN", ErrNumber, ErrDescription, Context)
    End Select
End Function


' **********************************************************************************************' log file
' **********************************************************************************************

Private Sub OpenLog(ByVal FilePath As String)
    LogFileNo = FreeFile
    Open FilePath For Append As #LogFileNo
End Sub


Private Sub WriteLog(ByVal Text As String)
    If LogFileNo = 0 Then Exit Sub
    Print #LogFileNo, Format(Now, "yyyy-mm-dd hh:nn:ss") & vbTab & Text
End Sub


Private Sub CloseLog()
    If LogFileNo = 0 Then Exit Sub
    Close #LogFileNo
    LogFileNo = 0
End Sub


' **********************************************************************************************
' additional controls
' **********************************************************************************************

Private Sub BuildExtendedControls()
    ' adds a panel right of the original controls; created at runtime because the form layout (.frx) is binary;
    ' the texts of all controls are set in ApplyLanguage
    Dim PanelLeft As Single
    Dim Fra As MSForms.Frame
    Dim i As Long

    ProgressBarWidth = 282
    PanelLeft = Me.InsideWidth + 6
    Me.Width = Me.Width + PanelWidth + 12

    ' language of the dialog
    Set lblOldSite = AddControl(Me, "Forms.Label.1", "ext_label_oldsite", 0, 0, 200, 12)
    Set lblLanguage = AddControl(Me, "Forms.Label.1", "ext_label_language", PanelLeft, 9, 80, 12)
    Set cboLanguage = AddControl(Me, "Forms.ComboBox.1", "ext_language", PanelLeft + 84, 6, 120, 18)
    cboLanguage.Style = fmStyleDropDownList
    cboLanguage.AddItem "English"
    cboLanguage.AddItem "Deutsch"
    cboLanguage.ListIndex = IIf(CurrentLanguage = LangGerman, 1, 0)

    ' additional options
    Set Fra = AddControl(Me, "Forms.Frame.1", "ext_frame_options", PanelLeft, PanelTop, PanelWidth, 136)
    Set fraOptions = Fra

    Set lblExisting = AddControl(Fra, "Forms.Label.1", "ext_label_existing", 6, 8, 140, 12)
    Set cboExisting = AddControl(Fra, "Forms.ComboBox.1", "ext_existing", 148, 5, 142, 16)
    cboExisting.Style = fmStyleDropDownList
    FillCombo cboExisting, "CBO_EX_ASK,CBO_EX_OVERWRITE,CBO_EX_SKIP,CBO_EX_NUMBER"

    Set chkLog = AddControl(Fra, "Forms.CheckBox.1", "ext_log", 6, 28, 284, 16)

    Set chkStopEmptyRecord = AddControl(Fra, "Forms.CheckBox.1", "ext_stop_empty_record", 6, 88, 284, 28)
    chkStopEmptyRecord.WordWrap = True

    Set chkAbortEmpty = AddControl(Fra, "Forms.CheckBox.1", "ext_abort_empty", 6, 48, 284, 16)
    Set cboAbortField = AddControl(Fra, "Forms.ComboBox.1", "ext_abort_field", 18, 66, 172, 16)
    cboAbortField.Style = fmStyleDropDownList
    For i = 1 To FieldCount
        cboAbortField.AddItem FieldNames(i)
    Next i

    ' preview of the file names
    Set fraPreview = AddControl(Me, "Forms.Frame.1", "ext_frame_preview", PanelLeft, PanelTop + 142, PanelWidth, 262)

    Set lblCount = AddControl(fraPreview, "Forms.Label.1", "ext_count", 6, 4, 200, 36)
    lblCount.WordWrap = True
    Set btnPreview = AddControl(fraPreview, "Forms.CommandButton.1", "ext_preview_refresh", 210, 4, 80, 20)

    Set chkQuickScan = AddControl(fraPreview, "Forms.CheckBox.1", "ext_quick_scan", 6, 42, 284, 16)
    If StorageQuickScan <> "" Then
        chkQuickScan.Value = (StorageQuickScan = "1")
    Else
        chkQuickScan.Value = DefaultQuickScan
    End If

    ' columns: record, file name, note; the columns have a fixed width and the list scrolls horizontally
    Set lstPreview = AddControl(fraPreview, "Forms.ListBox.1", "ext_preview_list", 6, 62, 284, 132)
    lstPreview.ColumnCount = 3
    lstPreview.ColumnWidths = PreviewMinColumnWidths

    ' shows the complete file name and note of the selected row of the list
    Set lblDetail = AddControl(fraPreview, "Forms.Label.1", "ext_preview_detail", 6, 198, 284, 44)
    lblDetail.WordWrap = True

    ' question for existing files, shown instead of the preview during the export
    Set fraConflict = AddControl(Me, "Forms.Frame.1", "ext_frame_conflict", PanelLeft, PanelTop + 142, PanelWidth, 140)
    fraConflict.Visible = False

    Set lblConflict = AddControl(fraConflict, "Forms.Label.1", "ext_conflict_text", 6, 6, 284, 60)
    lblConflict.WordWrap = True
    Set btnConflictOverwrite = AddControl(fraConflict, "Forms.CommandButton.1", "ext_conflict_overwrite", 6, 72, 66, 20)
    Set btnConflictSkip = AddControl(fraConflict, "Forms.CommandButton.1", "ext_conflict_skip", 78, 72, 66, 20)
    Set btnConflictNumber = AddControl(fraConflict, "Forms.CommandButton.1", "ext_conflict_number", 150, 72, 66, 20)
    Set btnConflictAbort = AddControl(fraConflict, "Forms.CommandButton.1", "ext_conflict_abort", 222, 72, 66, 20)
    Set chkConflictAll = AddControl(fraConflict, "Forms.CheckBox.1", "ext_conflict_all", 6, 100, 284, 16)

    ' progress with the remaining time
    Set Fra = AddControl(Me, "Forms.Frame.1", "ext_frame_progress", PanelLeft, PanelTop + 410, PanelWidth, 68)
    Set fraProgress = Fra

    Set lblProgress = AddControl(Fra, "Forms.Label.1", "ext_progress_text", 6, 4, 284, 12)
    Set lblProgressBack = AddControl(Fra, "Forms.Label.1", "ext_progress_back", 6, 20, ProgressBarWidth + 2, 12)
    lblProgressBack.BackStyle = fmBackStyleOpaque
    lblProgressBack.BackColor = vbWhite
    lblProgressBack.SpecialEffect = fmSpecialEffectSunken
    Set lblProgressBar = AddControl(Fra, "Forms.Label.1", "ext_progress_bar", 7, 21, 1, 10)
    lblProgressBar.BackStyle = fmBackStyleOpaque
    lblProgressBar.BackColor = RGB(0, 120, 215)
    lblProgressBar.Visible = False
    Set lblEta = AddControl(Fra, "Forms.Label.1", "ext_progress_eta", 6, 36, 284, 12)

    ' make the form high enough for the panel
    If Me.InsideHeight < 516 Then Me.Height = Me.Height + (516 - Me.InsideHeight)
End Sub

Private Function AddControl(Container As Object, ByVal ProgID As String, ByVal CtlName As String, ByVal PosLeft As Single, ByVal PosTop As Single, ByVal CtlWidth As Single, ByVal CtlHeight As Single) As Object
    Dim Ctl As Object

    Set Ctl = Container.Controls.Add(ProgID, CtlName, True)
    Ctl.Left = PosLeft
    Ctl.Top = PosTop
    Ctl.Width = CtlWidth
    Ctl.Height = CtlHeight
    If ProgID = "Forms.Label.1" Or ProgID = "Forms.CheckBox.1" Then Ctl.Caption = ""
    Set AddControl = Ctl
End Function


' **********************************************************************************************
' helpers
' **********************************************************************************************

Private Function SelectListItem(Box As MSForms.ComboBox, ByVal Text As String) As Boolean
    Dim i As Long

    For i = 0 To Box.ListCount - 1
        If StrComp(Box.List(i), Text, vbTextCompare) = 0 Then
            Box.ListIndex = i
            SelectListItem = True
            Exit Function
        End If
    Next i
End Function


Private Function IsMac() As Boolean
    IsMac = (System.OperatingSystem Like "*Mac*")
End Function


Private Function IsNewMac() As Boolean
    IsNewMac = IsMac() And Val(Application.Version) > 15
End Function
