VERSION 5.00
Object = "{48E59290-9880-11CF-9754-00AA00C00908}#1.0#0"; "MSINET.OCX"
Begin VB.Form Form1 
   Caption         =   "Auto update"
   ClientHeight    =   3000
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   8100
   LinkTopic       =   "Form1"
   ScaleHeight     =   3000
   ScaleWidth      =   8100
   StartUpPosition =   3  'Windows Default
   Begin InetCtlsObjects.Inet xHTTP 
      Left            =   4080
      Top             =   120
      _ExtentX        =   1005
      _ExtentY        =   1005
      _Version        =   393216
   End
   Begin VB.ListBox List1 
      BeginProperty Font 
         Name            =   "Lucida Console"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2205
      Left            =   120
      TabIndex        =   1
      Top             =   720
      Width           =   7935
   End
   Begin VB.Label lblStatus 
      Height          =   495
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7935
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private FFile As Integer
Private strCurrentAction As String

Private Function GetFile(strPath As String, ToFile As String) As String

    Dim bData() As Byte
    Dim GF As Variant
    Dim iTry As Integer
    iTry = 0
    
    On Error GoTo ErrHandler
    
    If strPath = "" Or ToFile = "" Then Exit Function
    
    DoEvents
    
    xHTTP.RequestTimeout = TimeOut
    bData() = xHTTP.OpenURL(strPath, icByteArray)
    DoEvents
    
    FFile = FreeFile
    Open ToFile For Binary Access Write As #FFile
    Put #FFile, , bData()
    Close #FFile
    
    GetFile = "OK"

    Exit Function
    
ErrHandler:
    
    Select Case Err.Number
    Case 35761 ' timeout
        GetFile = "TMO"
        iTry = iTry + 1
        If iTry <= Retry Then
            lblStatus.Caption = "Try " & (iTry + 1)
            Resume
        End If
    Case Else
        MsgBox "An unexpected error occured." & vbCrLf & Err.Number & vbCrLf & Err.Description, vbCritical, "Auto Update"
        GetFile = "NOK"
    End Select

End Function

Private Sub Form_Load()

    Dim RV As String

    Me.Show
    DoEvents
    
    lblStatus.Caption = "Retrieving version file..."
    
    On Error Resume Next
    Kill App.Path & "\versions.dat"
    On Error GoTo 0
    RV = GetFile(UpdateFile, App.Path & "\versions.dat")
    
    If RV = "OK" Then
        
        Dim MFFile As Integer
        Dim LocFile As Integer
        
        MFFile = FreeFile
        Open App.Path & "\versions.dat" For Input As #MFFile
                
        Dim FileName As String
        Dim FileDate As Date
        Dim Destination As String
                
        Do Until EOF(MFFile)
            Input #MFFile, FileName, Destination, FileDate
            Destination = Replace(Destination, "##apppath##", App.Path, , , vbTextCompare)
            List1.AddItem "BUS " & Destination
            List1.ListIndex = List1.ListCount - 1
            lblStatus.Caption = "Updating " & Destination
            DoEvents
            If IsOld(Destination, FileDate) Then
                ' wait a few moments
                ' prevents timeouts
                For t = 1 To 10
                    Sleep 100
                    DoEvents
                Next t
                RV = GetFile(FileName, Destination)
                If RV = "OK" Then
                    ' succeeded
                    List1.List(List1.ListCount - 1) = "OK  " & Destination
                ElseIf RV = "TMO" Then
                    ' timed out
                    List1.List(List1.ListCount - 1) = "TMO " & Destination
                Else
                    ' failed
                    List1.List(List1.ListCount - 1) = "FLD " & Destination
                End If
            Else
                ' no need to update
                List1.List(List1.ListCount - 1) = "SKP " & Destination
            End If
            DoEvents
        Loop
        
        Close #MFFile
        
        If AppToRun <> "" Then
            lblStatus.Caption = "Starting application..."
            
            On Error Resume Next
            Shell App.Path & "\" & AppToRun, vbNormalFocus
        End If
        
        ' wait 5 seconds before closing
        
        For t = 0 To 50
            Sleep 100
            DoEvents
        Next t
        
        End
    Else
        MsgBox "Cannot obtain version file", vbCritical, "Auto Update"
        On Error Resume Next
        ' try to start the specified application
        Shell App.Path & "\" & AppToRun, vbNormalFocus
        End
    End If

End Sub
