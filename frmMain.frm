VERSION 5.00
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Object = "{65E121D4-0C60-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCHRT20.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "Mortgage Calculator"
   ClientHeight    =   8595
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   11880
   LinkTopic       =   "Form1"
   ScaleHeight     =   8595
   ScaleWidth      =   11880
   StartUpPosition =   2  'CenterScreen
   WindowState     =   2  'Maximized
   Begin VB.CommandButton cmdCopyChart 
      Caption         =   "&Copy Chart"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   1320
      TabIndex        =   13
      Top             =   3600
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog cdlExport 
      Left            =   3000
      Top             =   2520
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.CommandButton cmdToHTML 
      Caption         =   "To &HTML"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   720
      TabIndex        =   11
      Top             =   3020
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "Graph"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   2640
      TabIndex        =   20
      Top             =   3120
      Visible         =   0   'False
      Width           =   1455
      Begin VB.OptionButton optGraph 
         Caption         =   "Line"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   15
         Top             =   600
         Width           =   1215
      End
      Begin VB.OptionButton optGraph 
         Caption         =   "Area"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   14
         Top             =   300
         Width           =   1215
      End
   End
   Begin MSChart20Lib.MSChart MSChart1 
      Height          =   3975
      Left            =   4200
      OleObjectBlob   =   "frmMain.frx":0000
      TabIndex        =   19
      Top             =   120
      Visible         =   0   'False
      Width           =   7575
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   3975
      Left            =   120
      TabIndex        =   18
      Top             =   4200
      Visible         =   0   'False
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   7011
      _Version        =   393216
      Cols            =   8
      FixedCols       =   0
      AllowUserResizing=   1
      FormatString    =   ""
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton cmdChart 
      Caption         =   "&View Chart"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   12
      Top             =   3600
      Width           =   1215
   End
   Begin VB.CommandButton cmdAmortise 
      Caption         =   "&Amortisation Table"
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   10
      Top             =   2520
      Width           =   2415
   End
   Begin VB.TextBox txtResult 
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   16
      Text            =   "0"
      Top             =   1980
      Width           =   1215
   End
   Begin VB.CommandButton cmdCalculate 
      Caption         =   "&Monthly Payment"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   120
      TabIndex        =   9
      Top             =   1920
      Width           =   2415
   End
   Begin MSComCtl2.DTPicker dtpStart 
      Height          =   375
      Left            =   2520
      TabIndex        =   8
      Top             =   1500
      Width           =   1575
      _ExtentX        =   2778
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      CustomFormat    =   "dd/MM/yyy"
      Format          =   24510467
      CurrentDate     =   36908
   End
   Begin VB.TextBox txtTerm 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   7
      Top             =   1020
      Width           =   735
   End
   Begin VB.TextBox txtInterest 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   5
      Top             =   540
      Width           =   735
   End
   Begin VB.TextBox txtLoan 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2520
      TabIndex        =   4
      Top             =   60
      Width           =   1575
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "="
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   240
      Left            =   2640
      TabIndex        =   17
      Top             =   2040
      Width           =   135
   End
   Begin VB.Label Label5 
      Caption         =   "%"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   3240
      TabIndex        =   6
      Top             =   600
      Width           =   255
   End
   Begin VB.Label Label4 
      Caption         =   "Date of First Payment:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   1560
      Width           =   2295
   End
   Begin VB.Label Label3 
      Caption         =   "Term (Whole Years):"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1080
      Width           =   2295
   End
   Begin VB.Label Label2 
      Caption         =   "Interest Rate:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   2295
   End
   Begin VB.Label Label1 
      Caption         =   "Loan Amount:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   2295
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private iPayments As Integer ' iPayments = no. of MONTHLY payments payable

Private Sub cmdAmortise_Click()

'This sub populates the MSFlexgrid

    On Error Resume Next ' to compensate for the error sometimes given in the cell colour changing routine
    Dim sCurrMonth As String 'for the Date column
    Dim iDay, iMonth, iYear, iCountPay As Integer ' 1st 3 are for the date column, 4th is the loop counter
    Dim dCurrBal, dCurrInt, dCurrPrin, dEndBal, dCumInt, dTotPaid As Double 'Period's starting loan balance, interest this period,
                                                                            'principal paid this period, loan balance at end of period,
                                                                            'cumulative interest at end of period, total paid to date at end of period
                                                                            '; respectively!
    Dim sWhen As String 'the commencement date
    
    frmMain.MousePointer = vbHourglass ' cpu busy
    sWhen = dtpStart ' assign commencement date
    iDay = Val(Left(sWhen, 2)) ' the day
    iMonth = Val(Mid(sWhen, 4, 2)) ' the month
    iYear = Val(Right(sWhen, 2)) ' the year - latter 2 are required to increment period
    dTotPaid = txtResult.Text ' assign monthly payment
    dCurrBal = txtLoan.Text ' assign loan balance for 1st period
    dCurrInt = dCurrBal / 12 * (txtInterest / 100) ' assign interest for the first period
    dCurrPrin = dTotPaid - dCurrInt ' assign principal paid for the first period
    dEndBal = dCurrBal - dCurrPrin ' assign balance at end of the first period
    dCumInt = dCurrInt ' assign cumulative interest for the first period
    dTotPaid = txtResult.Text ' assign running total paid for first period
    
    With MSFlexGrid1
        .Rows = iPayments + 1 ' number of rows in the FlexGrid. + 1, to cater for the header row
        .TextMatrix(1, 2) = Format(dCurrBal, "Currency", 2)  ' print current balance,
        .TextMatrix(1, 3) = Format(dCurrInt, "Currency", 2)  ' current interest,
        .TextMatrix(1, 4) = Format(dCurrPrin, "Currency", 2) ' current principal,
        .TextMatrix(1, 5) = Format(dEndBal, "Currency", 2)   ' end balance,
        .TextMatrix(1, 6) = Format(dCumInt, "Currency", 2)   ' cumulative interest, and
        .TextMatrix(1, 7) = Format(dTotPaid, "Currency", 2)  ' running total paid
    End With                                                 ' for the first period.
                                                             ' Remeber that the FlexGrid row and column indices both begin at 0.
                                                             ' So the current balance will print in the 3rd column from the left,
                                                             ' the current interest will appear in the 4th, and so on.
    ' the all-important loop
    For iCountPay = 1 To iPayments ' 1 to the number of rows
        With MSFlexGrid1
            .TextMatrix(iCountPay, 0) = iCountPay ' populates the No. of payments column i.e. the first column on the left
            If iMonth > 12 Then   ' If we're past December, we revert
                iMonth = 1        ' back to January,
                iYear = iYear + 1 ' and add a year.
            End If
            .TextMatrix(iCountPay, 1) = Format(Val(iDay), "00") & _
            "/" & Format(Val(iMonth), "00") & "/" & Format(Val(iYear), "00")
            ' The above split line populates the 2nd column with the payment dates.
            ' I get the feeling I could have done this more simply...
            ' Basically, I converted the bits of the month back to little strings,
            ' and concatenated them with slashes to form 1 big, lovely string.
            iMonth = iMonth + 1 ' increment the month
            If iCountPay > 1 Then ' the loop for the financial figures really begins
                                  ' from the 3rd row (1st row=headers, 2nd was the first period,
                                  ' and was already populated above.
                
                ' starting balance for this period is read from ending balance from previous period
                dCurrBal = .TextMatrix(iCountPay - 1, 5)
                ' print the starting balance for this period
                .TextMatrix(iCountPay, 2) = Format(dCurrBal, "Currency", 2)
                ' multiply the starting balance by the interst rate (which is divided by 100 to provide a percentage)
                ' and divide by 12 to get the interest paid for this month.
                dCurrInt = .TextMatrix(iCountPay, 2) / 12 * (txtInterest / 100)
                ' print the interest paid this period
                .TextMatrix(iCountPay, 3) = Format(dCurrInt, "Currency", 2)
                ' this periods principal paid is the monthly total minus the interest paid
                dCurrPrin = txtResult - dCurrInt
                ' print the principal paid for this period
                .TextMatrix(iCountPay, 4) = Format(dCurrPrin, "Currency", 2)
                ' this periods end balance is the start balance minus principal paid this period
                dEndBal = dCurrBal - dCurrPrin
                ' print the end balance for this period
                .TextMatrix(iCountPay, 5) = Format(dEndBal, "Currency", 2)
                ' cumulative interest for this period is the cumulative interst paid in the last period
                ' plus interest paid this period
                dCumInt = .TextMatrix(iCountPay - 1, 6) + dCurrInt
                ' print the cumulative interest for this period
                .TextMatrix(iCountPay, 6) = Format(dCumInt, "Currency", 2)
                ' the total paid to date this period is simply the monthly payment
                ' multiplied by the payment number
                dTotPaid = .TextMatrix(1, 7) * iCountPay
                ' print the total paid at the end of this period
                .TextMatrix(iCountPay, 7) = Format(dTotPaid, "Currency", 2)
            End If
        End With
    Next iCountPay ' increment the all-important loop
    
    With MSFlexGrid1
        .ColWidth(0) = 640   ' Adjust the column widths
        .ColWidth(1) = 1220  ' in 'twips'.
        .ColWidth(2) = 1700
        .ColWidth(3) = 1400
        .ColWidth(4) = 1400
        .ColWidth(5) = 1700
        .ColWidth(6) = 1700
        .ColWidth(7) = 1700
        Dim i
        For i = 0 To 7           ' All columns
            .ColAlignment(i) = 3 ' to be centrally aligned
        Next i
    End With
    
    ' This section colours alternate rows, making the
    ' grid a little easier to read.
    Dim iCols As Integer
     ' for each row (select by row first)
    Do Until MSFlexGrid1.Row = iPayments
        ' skip the header row & select the next.
        ' start the first row white,
        ' and increment
        MSFlexGrid1.Row = MSFlexGrid1.Row + 1
        ' for all columns (remember that the Columns index
        ' begins with 0, but that the .Col property starts at 1
        ' hence the - 1
        For iCols = 0 To 7 'MSFlexGrid1.Cols - 1
            ' select each column in turn, after having already
            ' selected the row above
            MSFlexGrid1.Col = iCols
            ' change the cell colour (light green, in this case)
            MSFlexGrid1.CellBackColor = &HC0FFC0
        Next iCols
        ' next column, please
        MSFlexGrid1.Row = MSFlexGrid1.Row + 1
    Loop
    
    MSFlexGrid1.Visible = True      ' show the grid
    cmdChart.Enabled = True         ' show the chart button
    cmdToHTML.Enabled = True        ' show export to HTML button
    frmMain.MousePointer = vbNormal ' revert mouse pointer
End Sub

Private Sub cmdCalculate_Click()
'This calculates the monthly payment
    
    On Error GoTo BadData ' If you forget to fill in a field
    Dim dRate, dLoan, dMonthlyPayment As Double ' The rate, the principal, the eventual result
    
    ' This is the monthly interest rate. The interest rate per
    ' period is required for VB's own 'Pmt' function.
    ' It's divided by 1200 (a) to get the monthly rate: 12; and
    ' (b) to get a percentage: 100
    dRate = txtInterest.Text / 1200
    ' assign the principal
    dLoan = txtLoan.Text
    ' calculate the ubiquitous iPayments variable, which
    ' is required in the 'Pmt' function, and helps determine
    ' the number of rows within the FlexGrid
    iPayments = txtTerm * 12
    ' The 'Pmt' function calculates the monthly payment
    dMonthlyPayment = Pmt(dRate, iPayments, -dLoan)
    ' assign the result to the relevant text box and format
    ' it as currency to 2 decimal places
    txtResult.Text = Format(dMonthlyPayment, "Currency", 2)
    ' show the print amortisation schedule button
    cmdAmortise.Enabled = True
    ' exit the sub, so the program doesn't continue on to
    ' the 'BadData' error-handling section
    Exit Sub

' What can I say? I'm a lazy swine. I've used the routine for
' handling all errors...
BadData:
    MsgBox "You have input your data incorrectly!", vbOKOnly + vbCritical, "Error"
    cmdAmortise.Enabled = False
    cmdChart.Enabled = False
    txtLoan.SetFocus
End Sub

Private Sub cmdChart_Click()
' this plots and shows the chart

    ' Charts can be thought of as graphic representations
    ' of a matrix of figures. And generally they are handled
    ' this way. A chart's data is based on a grid. You enter
    ' data in rows and columns, thereby affecting the chart's
    ' appearance.
    Dim iRowCount, iColumnCount As Integer ' the chart's rows and columns
    Dim dCurrData As Double ' I'll assign a chart 'cells' data with this variable
    'Dim sDateLabel As String
    
    frmMain.MousePointer = vbHourglass ' cpu busy
    ' This ensures that the chart will not redraw itself
    ' every time data is placed within a cell using the
    ' loop below.
    MSChart1.Repaint = False
    MSChart1.RowCount = iPayments ' See? Just like the FlexGrid (well alomst)
    
    ' You can either select columns or rows to begin this
    ' little nested loop. I chose columns in this case.
    ' I don't know why...
    ' Beginning first loop
    ' Only 2 columns needed - each periods interest & principal payments
    For iColumnCount = 1 To 2
        ' You must also plot each row
        For iRowCount = 1 To iPayments
            ' assign a FlexGrid cell's data to dCurrData
            ' I use iColumnCount + 2, because the first loop
            ' will seek for data in column number 4 (interst),
            ' then column 5 (principal) - again remember that
            ' a FlexGrid's column and row indices begin at 0,
            ' so you have to compensate...
            dCurrData = MSFlexGrid1.TextMatrix(iRowCount, iColumnCount + 2)
            ' MSChart's SetData method allocates dCurrData to
            ' the chart's current cell
            MSChart1.DataGrid.SetData iRowCount, iColumnCount, dCurrData, False
        Next iRowCount
    Next iColumnCount
    
    ' These are the series' legends
    ' I changed these lines in this version so that the chart
    ' lables would paste correctly into a graphics package
    MSChart1.Column = 1
    MSChart1.ColumnLabel = "Interest"
    MSChart1.Column = 2
    MSChart1.ColumnLabel = "Principal"
    
    ' show the chart
    MSChart1.Visible = True
    ' repaint the chart
    MSChart1.Repaint = True
    ' show chart options
    Frame1.Visible = True
    ' clicking the chart will now not produce any ugle lines:
    MSChart1.Enabled = False
    cmdCopyChart.Enabled = True
    frmMain.MousePointer = vbNormal 'cpu finished
End Sub

Private Sub cmdCopyChart_Click()
    'well, this is embarrassingly easy...
    'you can now paste the chart into a graphics
    'package, convert it into a JPEG or GIF (you
    'can't do this with regular VB controls and manually
    'insert it into the HTML doc
    
    'alternatively, get Joe Oliphant's multi-format graphics
    'viewer/editor at:
    'http://www.winsite.com/info/pc/win95/programr/vbasic/gvocx.zip/index.html
    'I haven't used it yet, but I think it has the functionality
    'of the MS Image or Picture controls.
    'anybody who has it can try EditPaste the chart into the
    'control, automatically save it and add it into the HTML doc
    
    MSChart1.EditCopy
End Sub

Private Sub cmdToHTML_Click()
    Dim fExportFile As String 'the system name of the file to which you're exporting
    Dim iFileNo As Integer    'the file you're going to which you'll write the data
    Dim iCountRow As Integer  'main loop counter - will top off at iPayments (the no of months)
    Dim iCountCol As Integer  'nested loop counter (for Flexgrid columns 0 through 7)
    Dim sDate As String       'help add to the <title> in HTML section
    Dim sCellColour As String 'to get same alternating row colour effect as in the Flexgrid
    
    On Error GoTo ErrorTrap   'for the Common Dialogue CancelError event
    
    sDate = Format$(Date, "dd-mm-yy") 'determines and formats the date for inclusion into the <title>
    Screen.MousePointer = vbHourglass 'CPU busy
    With cdlExport 'the common dialogue control
        .CancelError = True 'when Cancel clicked prog will branch to error handling routine
        .FileName = "amortschedule.html" 'default filename - remove if you don't want this
        .DialogTitle = "Export to Web Page..." 'text in Common Dialogue's title bar
        .Filter = "Web Pages (*.HTML)|*.HTML" 'HTML files viewable only
        .DefaultExt = "HTML" 'save files with HTML extension by default
        .Flags = cdlOFNHideReadOnly Or cdlOFNOverwritePrompt 'first flag hides the read-only box, the
                                                             'second prompts user to overwrite existing file
        .ShowSave 'we're saving a file, as opposed to loading one
    End With 'use 'With' to reduce the '.'s - the fewer of these, the faster the prog
    
    fExportFile = cdlExport.FileName 'assign writable filename to file user named
    iFileNo = FreeFile() 'find a file to 'Open'
    Open fExportFile For Output As #iFileNo 'our named file will now be written to
    
    'Here we go! This assumes you know HTML, of course!
    'I have tried to neatly space the code, so it can be
    'edited further within a HTML editor
    Print #iFileNo, "<html>"
    Print #iFileNo, ""
    Print #iFileNo, "<head>"
    Print #iFileNo, ""
    Print #iFileNo, "<title>Loan Amortisation Schedule as at " & sDate & "</title>" 'add the date to the title!
    Print #iFileNo, ""
    Print #iFileNo, "<br>"
    'show details upon which the schedule is based
    Print #iFileNo, "Loan Amount: <b>" & Format(txtLoan, "Currency", 2) & "</b><br>"
    Print #iFileNo, "Interest Rate: <b>" & txtInterest & "%</b><br>"
    Print #iFileNo, "Loan Term: <b>" & txtTerm & " years</b><br>"
    Print #iFileNo, "Monthly Payment: <b>" & Format(txtResult, "Currency", 2) & "</b><br>"
    Print #iFileNo, "<br>"
    'My habit is to enclose tag attribute values within double quotation marks ("),
    'as is recommended by the W3C. This poses problems, however - strings will be
    'cut off if you type a ", and you'll be given a syntax error if you try to complete
    'a line with further "s. You can get around this by concatenating the string
    'with 'Chr(34)', as I've done below. You'll have to do this each time a " occurs.
    '34 is the ASCII number for ". This is also useful for entering certain formulae
    'within Excel if you're using VB/VBA.
    Print #iFileNo, "<table border=" & Chr(34) & "1" & Chr(34) & ">"
    Print #iFileNo, "<tr>"
    'the first row is always the same - headers!
    Print #iFileNo, "   <td align=" & Chr(34) & "center" & Chr(34) & " bgcolor=" & Chr(34) & "#999999" & Chr(34) & "><b>&nbsp;&nbsp;No.&nbsp;&nbsp;</b></td>"
    Print #iFileNo, "   <td align=" & Chr(34) & "center" & Chr(34) & " bgcolor=" & Chr(34) & "#999999" & Chr(34) & "><b>&nbsp;&nbsp;Date&nbsp;&nbsp;</b></td>"
    Print #iFileNo, "   <td align=" & Chr(34) & "center" & Chr(34) & " bgcolor=" & Chr(34) & "#999999" & Chr(34) & "><b>&nbsp;&nbsp;Begin Bal.&nbsp;&nbsp;</b></td>"
    Print #iFileNo, "   <td align=" & Chr(34) & "center" & Chr(34) & " bgcolor=" & Chr(34) & "#999999" & Chr(34) & "><b>&nbsp;&nbsp;Interest&nbsp;&nbsp;</b></td>"
    Print #iFileNo, "   <td align=" & Chr(34) & "center" & Chr(34) & " bgcolor=" & Chr(34) & "#999999" & Chr(34) & "><b>&nbsp;&nbsp;Principal&nbsp;&nbsp;</b></td>"
    Print #iFileNo, "   <td align=" & Chr(34) & "center" & Chr(34) & " bgcolor=" & Chr(34) & "#999999" & Chr(34) & "><b>&nbsp;&nbsp;End Bal.&nbsp;&nbsp;</b></td>"
    Print #iFileNo, "   <td align=" & Chr(34) & "center" & Chr(34) & " bgcolor=" & Chr(34) & "#999999" & Chr(34) & "><b>&nbsp;&nbsp;Cum. Interest&nbsp;&nbsp;</b></td>"
    Print #iFileNo, "   <td align=" & Chr(34) & "center" & Chr(34) & " bgcolor=" & Chr(34) & "#999999" & Chr(34) & "><b>&nbsp;&nbsp;Total Paid&nbsp;&nbsp;</b></td>"
    Print #iFileNo, "</tr>"
    
    For iCountRow = 1 To iPayments   'for each payment
        If iCountRow Mod 2 <> 0 Then 'if the rownumber is odd...
            sCellColour = "#ffffff"  'the row colour is white
        Else                         'otherwise...
            sCellColour = "#ccffcc"  'it's light green
        End If
        Print #iFileNo, "<tr>" 'start the HTML row for this iCOuntRow
        For iCountCol = 0 To 7 'start the nested loop for each column. This assumes
                               'all columns will be formatted the same way
                               
            'the next line assigns this cell's contents
            'note that the cell colour is determined by 'sCellColour'
            'lotsa Chr(34)s in here!
            Print #iFileNo, "   <td bgcolor=" & Chr(34) & sCellColour & Chr(34) & _
            "align=" & Chr(34) & "center" & Chr(34) & ">" & _
                MSFlexGrid1.TextMatrix(iCountRow, iCountCol) & "</td>"
        Next iCountCol 'increment baby column loop
        Print #iFileNo, "</tr>" 'close off the row in HTML
    Next iCountRow 'increment the mommy loop
    
    'closing off tags
    Print #iFileNo, "</table>"
    Print #iFileNo, ""
    Print #iFileNo, "</head>"
    Print #iFileNo, ""
    Print #iFileNo, "</html>"
    
    Close #iFileNo 'ALWAYS close an opened file...
    Screen.MousePointer = vbNormal 'CPU free
    'successful completion message box
    MsgBox "Export to HTML complete!", vbOKOnly + vbInformation, "Export Complete"
    Exit Sub
    
ErrorTrap:
    If Err.Number = 32755 Then '32755 = if the user clicked Cancel
        Screen.MousePointer = vbNormal 'CPU free
        Exit Sub 'then quit
    End If
    'but if, for some reason something else happened...
    MsgBox "Error encountered in exporting.", vbCritical, "Error"
    Screen.MousePointer = vbNormal 'CPU free
End Sub

Private Sub Form_Load()
    ' the FlexGrid's column headers...
    With MSFlexGrid1
        .TextMatrix(0, 0) = "No."
        .TextMatrix(0, 1) = "Date"
        .TextMatrix(0, 2) = "Begin Bal."
        .TextMatrix(0, 3) = "Interest"
        .TextMatrix(0, 4) = "Principal"
        .TextMatrix(0, 5) = "End Bal."
        .TextMatrix(0, 6) = "Cum. Interest"
        .TextMatrix(0, 7) = "Total Paid"
    End With
    ' assign today's date to the DateTimePicker control
    dtpStart = Date
    ' start off with an Area type graph
    optGraph(0).Value = True
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmMain = Nothing
End Sub

Private Sub optGraph_Click(Index As Integer)
' sub to switch graph types
    If Index = 0 Then
        MSChart1.chartType = VtChChartType2dArea
        ' one series appears on top of the other
        MSChart1.Stacking = True
        ' redraw the chart
        MSChart1.Repaint = True
    Else
        MSChart1.chartType = VtChChartType2dLine
        ' series shown apart - necessary for a meaningful line graph
        MSChart1.Stacking = False
        ' redraw the chart
        MSChart1.Repaint = True
    End If
End Sub

Private Sub txtInterest_Change()
    Call Reset ' always use the 'call' keyword when referring
               ' to your own subs - helps avoid confusion
End Sub

Private Sub txtLoan_Change()
    Call Reset ' always use the 'call' keyword when referring
               ' to your own subs - helps avoid confusion
End Sub

Private Sub txtTerm_Change()
    Call Reset ' always use the 'call' keyword when referring
               ' to your own subs - helps avoid confusion
End Sub

Sub Reset()
' altering any inputs will reset results
' I don't know why I didn't have this as 1 sub in the
' first version...!
    txtResult.Text = 0
    cmdAmortise.Enabled = False
    cmdChart.Enabled = False
    cmdToHTML.Enabled = False
    cmdCopyChart.Enabled = False
    MSFlexGrid1.Visible = False
    MSChart1.Visible = False
End Sub
