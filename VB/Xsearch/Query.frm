VERSION 5.00
Object = "{6B7E6392-850A-101B-AFC0-4210102A8DA7}#1.3#0"; "comctl32.ocx"
Begin VB.Form frmQuery 
   Caption         =   "Cross Library Search"
   ClientHeight    =   7920
   ClientLeft      =   60
   ClientTop       =   348
   ClientWidth     =   11376
   LinkTopic       =   "Form1"
   ScaleHeight     =   7920
   ScaleWidth      =   11376
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame3 
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      Height          =   4545
      Left            =   240
      TabIndex        =   47
      Top             =   420
      Width           =   9345
      Begin VB.ComboBox cboLib2Select 
         Height          =   315
         Index           =   9
         Left            =   5160
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   67
         Top             =   4200
         Width           =   3135
      End
      Begin VB.ComboBox cboLib2Select 
         Height          =   315
         Index           =   8
         Left            =   5160
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   65
         Top             =   3840
         Width           =   3135
      End
      Begin VB.ComboBox cboLib2Select 
         Height          =   315
         Index           =   7
         Left            =   5160
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   63
         Top             =   3480
         Width           =   3135
      End
      Begin VB.ComboBox cboLib2Select 
         Height          =   315
         Index           =   6
         Left            =   5160
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   61
         Top             =   3120
         Width           =   3135
      End
      Begin VB.ComboBox cboLib2Select 
         Height          =   315
         Index           =   5
         Left            =   5160
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   59
         Top             =   2760
         Width           =   3135
      End
      Begin VB.ComboBox cboLib2Select 
         Height          =   315
         Index           =   4
         Left            =   5160
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   57
         Top             =   2400
         Width           =   3135
      End
      Begin VB.ComboBox cboLib2Select 
         Height          =   315
         Index           =   3
         Left            =   5160
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   55
         Top             =   2040
         Width           =   3135
      End
      Begin VB.ComboBox cboLib2Select 
         Height          =   315
         Index           =   2
         Left            =   5160
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   53
         Top             =   1680
         Width           =   3135
      End
      Begin VB.ComboBox cboLib2Select 
         Height          =   315
         Index           =   1
         Left            =   5160
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   51
         Top             =   1320
         Width           =   3135
      End
      Begin VB.ComboBox cboLib2Select 
         Height          =   315
         Index           =   0
         Left            =   5160
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   49
         Top             =   960
         Width           =   3135
      End
      Begin VB.ComboBox cboLib1Select 
         Height          =   315
         Index           =   9
         Left            =   1080
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   66
         Top             =   4200
         Width           =   3135
      End
      Begin VB.ComboBox cboLib1Select 
         Height          =   315
         Index           =   8
         Left            =   1080
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   64
         Top             =   3840
         Width           =   3135
      End
      Begin VB.ComboBox cboLib1Select 
         Height          =   315
         Index           =   7
         Left            =   1080
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   62
         Top             =   3480
         Width           =   3135
      End
      Begin VB.ComboBox cboLib1Select 
         Height          =   315
         Index           =   6
         Left            =   1080
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   60
         Top             =   3120
         Width           =   3135
      End
      Begin VB.ComboBox cboLib1Select 
         Height          =   315
         Index           =   5
         Left            =   1080
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   58
         Top             =   2760
         Width           =   3135
      End
      Begin VB.ComboBox cboLib1Select 
         Height          =   315
         Index           =   4
         Left            =   1080
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   56
         Top             =   2400
         Width           =   3135
      End
      Begin VB.ComboBox cboLib1Select 
         Height          =   315
         Index           =   3
         Left            =   1080
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   54
         Top             =   2040
         Width           =   3135
      End
      Begin VB.ComboBox cboLib1Select 
         Height          =   315
         Index           =   2
         Left            =   1080
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   52
         Top             =   1680
         Width           =   3135
      End
      Begin VB.ComboBox cboLib1Select 
         Height          =   315
         Index           =   1
         Left            =   1080
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   50
         Top             =   1320
         Width           =   3135
      End
      Begin VB.ComboBox cboLib1Select 
         Height          =   315
         Index           =   0
         Left            =   1080
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   48
         Top             =   960
         Width           =   3135
      End
      Begin VB.Frame Frame6 
         Caption         =   "Libraries"
         Height          =   735
         Left            =   120
         TabIndex        =   83
         Top             =   0
         Width           =   9015
         Begin VB.Label lblLib2 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   5160
            TabIndex        =   85
            Top             =   240
            Width           =   3135
         End
         Begin VB.Label lblLib1 
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   255
            Left            =   960
            TabIndex        =   84
            Top             =   240
            Width           =   3135
         End
      End
   End
   Begin VB.Frame Frame2 
      BorderStyle     =   0  'None
      Caption         =   "Frame2"
      Height          =   4545
      Left            =   240
      TabIndex        =   24
      Top             =   420
      Width           =   9345
      Begin VB.ComboBox cboSearchField 
         Height          =   315
         Index           =   5
         Left            =   1080
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   44
         Top             =   2400
         Width           =   3135
      End
      Begin VB.ComboBox cboLogOp 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   5
         Left            =   4320
         Style           =   2  'Dropdown List
         TabIndex        =   45
         Top             =   2400
         Width           =   855
      End
      Begin VB.TextBox txtSearchData 
         Height          =   285
         Index           =   5
         Left            =   5280
         TabIndex        =   46
         Top             =   2400
         Width           =   3135
      End
      Begin VB.CommandButton cmdRelOp 
         Caption         =   "AND"
         Height          =   255
         Index           =   4
         Left            =   480
         TabIndex        =   43
         Top             =   2400
         Width           =   495
      End
      Begin VB.Frame Frame5 
         Caption         =   "Key"
         Height          =   735
         Left            =   120
         TabIndex        =   82
         Top             =   0
         Width           =   9015
         Begin VB.ComboBox cboSearchField 
            Height          =   315
            Index           =   0
            Left            =   960
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   25
            Top             =   240
            Width           =   3135
         End
         Begin VB.ComboBox cboLogOp 
            BackColor       =   &H8000000F&
            BeginProperty Font 
               Name            =   "MS Sans Serif"
               Size            =   7.8
               Charset         =   0
               Weight          =   700
               Underline       =   0   'False
               Italic          =   0   'False
               Strikethrough   =   0   'False
            EndProperty
            Height          =   315
            Index           =   0
            Left            =   4200
            Style           =   2  'Dropdown List
            TabIndex        =   26
            Top             =   240
            Width           =   855
         End
         Begin VB.TextBox txtSearchData 
            Height          =   285
            Index           =   0
            Left            =   5160
            TabIndex        =   27
            Top             =   240
            Width           =   3135
         End
      End
      Begin VB.CommandButton cmdRelOp 
         Caption         =   "AND"
         Height          =   255
         Index           =   2
         Left            =   480
         TabIndex        =   39
         Top             =   2040
         Width           =   495
      End
      Begin VB.CommandButton cmdRelOp 
         Caption         =   "AND"
         Height          =   255
         Index           =   1
         Left            =   480
         TabIndex        =   35
         Top             =   1680
         Width           =   495
      End
      Begin VB.CommandButton cmdRelOp 
         Caption         =   "AND"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   31
         Top             =   1320
         Width           =   495
      End
      Begin VB.TextBox txtSearchData 
         Height          =   285
         Index           =   4
         Left            =   5280
         TabIndex        =   42
         Top             =   2040
         Width           =   3135
      End
      Begin VB.TextBox txtSearchData 
         Height          =   285
         Index           =   3
         Left            =   5280
         TabIndex        =   38
         Top             =   1680
         Width           =   3135
      End
      Begin VB.TextBox txtSearchData 
         Height          =   285
         Index           =   2
         Left            =   5280
         TabIndex        =   34
         Top             =   1320
         Width           =   3135
      End
      Begin VB.TextBox txtSearchData 
         Height          =   285
         Index           =   1
         Left            =   5280
         TabIndex        =   30
         Top             =   960
         Width           =   3135
      End
      Begin VB.ComboBox cboLogOp 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   4
         Left            =   4320
         Style           =   2  'Dropdown List
         TabIndex        =   41
         Top             =   2040
         Width           =   855
      End
      Begin VB.ComboBox cboLogOp 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   3
         Left            =   4320
         Style           =   2  'Dropdown List
         TabIndex        =   37
         Top             =   1680
         Width           =   855
      End
      Begin VB.ComboBox cboLogOp 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   2
         Left            =   4320
         Style           =   2  'Dropdown List
         TabIndex        =   33
         Top             =   1320
         Width           =   855
      End
      Begin VB.ComboBox cboLogOp 
         BackColor       =   &H8000000F&
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   7.8
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   315
         Index           =   1
         Left            =   4320
         Style           =   2  'Dropdown List
         TabIndex        =   29
         Top             =   960
         Width           =   855
      End
      Begin VB.ComboBox cboSearchField 
         Height          =   315
         Index           =   4
         Left            =   1080
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   40
         Top             =   2040
         Width           =   3135
      End
      Begin VB.ComboBox cboSearchField 
         Height          =   315
         Index           =   3
         Left            =   1080
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   36
         Top             =   1680
         Width           =   3135
      End
      Begin VB.ComboBox cboSearchField 
         Height          =   315
         Index           =   2
         Left            =   1080
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   32
         Top             =   1320
         Width           =   3135
      End
      Begin VB.ComboBox cboSearchField 
         Height          =   315
         Index           =   1
         Left            =   1080
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   28
         Top             =   960
         Width           =   3135
      End
   End
   Begin VB.Frame Frame1 
      BorderStyle     =   0  'None
      Caption         =   "Frame1"
      Height          =   4545
      Left            =   240
      TabIndex        =   0
      Top             =   420
      Width           =   9345
      Begin VB.ComboBox cboLib2XRef 
         Height          =   315
         Index           =   5
         Left            =   5280
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   15
         Top             =   2760
         Width           =   3135
      End
      Begin VB.ComboBox cboLib2XRef 
         Height          =   315
         Index           =   6
         Left            =   5280
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   17
         Top             =   3120
         Width           =   3135
      End
      Begin VB.ComboBox cboLib2XRef 
         Height          =   315
         Index           =   7
         Left            =   5280
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   19
         Top             =   3480
         Width           =   3135
      End
      Begin VB.ComboBox cboLib2XRef 
         Height          =   315
         Index           =   8
         Left            =   5280
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   21
         Top             =   3840
         Width           =   3135
      End
      Begin VB.ComboBox cboLib2XRef 
         Height          =   315
         Index           =   9
         Left            =   5280
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   23
         Top             =   4200
         Width           =   3135
      End
      Begin VB.ComboBox cboLib1XRef 
         Height          =   315
         Index           =   5
         Left            =   1080
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   14
         Top             =   2760
         Width           =   3135
      End
      Begin VB.ComboBox cboLib1XRef 
         Height          =   315
         Index           =   6
         Left            =   1080
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   16
         Top             =   3120
         Width           =   3135
      End
      Begin VB.ComboBox cboLib1XRef 
         Height          =   315
         Index           =   7
         Left            =   1080
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   18
         Top             =   3480
         Width           =   3135
      End
      Begin VB.ComboBox cboLib1XRef 
         Height          =   315
         Index           =   8
         Left            =   1080
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   20
         Top             =   3840
         Width           =   3135
      End
      Begin VB.ComboBox cboLib1XRef 
         Height          =   315
         Index           =   9
         Left            =   1080
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   22
         Top             =   4200
         Width           =   3135
      End
      Begin VB.Frame Frame4 
         Caption         =   "Libraries"
         Height          =   735
         Left            =   120
         TabIndex        =   76
         Top             =   0
         Width           =   9015
         Begin VB.ComboBox cboLib2 
            Height          =   315
            Left            =   5160
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   3
            Top             =   240
            Width           =   3135
         End
         Begin VB.ComboBox cboLib1 
            Height          =   315
            Left            =   960
            Sorted          =   -1  'True
            Style           =   2  'Dropdown List
            TabIndex        =   2
            Top             =   240
            Width           =   3135
         End
      End
      Begin VB.ComboBox cboLib2XRef 
         Height          =   315
         Index           =   4
         Left            =   5280
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   13
         Top             =   2400
         Width           =   3135
      End
      Begin VB.ComboBox cboLib2XRef 
         Height          =   315
         Index           =   3
         Left            =   5280
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   2040
         Width           =   3135
      End
      Begin VB.ComboBox cboLib2XRef 
         Height          =   315
         Index           =   2
         Left            =   5280
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   9
         Top             =   1680
         Width           =   3135
      End
      Begin VB.ComboBox cboLib2XRef 
         Height          =   315
         Index           =   1
         Left            =   5280
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   7
         Top             =   1320
         Width           =   3135
      End
      Begin VB.ComboBox cboLib2XRef 
         Height          =   315
         Index           =   0
         Left            =   5280
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   5
         Top             =   960
         Width           =   3135
      End
      Begin VB.ComboBox cboLib1XRef 
         Height          =   315
         Index           =   4
         Left            =   1080
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   12
         Top             =   2400
         Width           =   3135
      End
      Begin VB.ComboBox cboLib1XRef 
         Height          =   315
         Index           =   3
         Left            =   1080
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   10
         Top             =   2040
         Width           =   3135
      End
      Begin VB.ComboBox cboLib1XRef 
         Height          =   315
         Index           =   2
         Left            =   1080
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   8
         Top             =   1680
         Width           =   3135
      End
      Begin VB.ComboBox cboLib1XRef 
         Height          =   315
         Index           =   1
         Left            =   1080
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   6
         Top             =   1320
         Width           =   3135
      End
      Begin VB.ComboBox cboLib1XRef 
         Height          =   315
         Index           =   0
         Left            =   1080
         Sorted          =   -1  'True
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   960
         Width           =   3135
      End
      Begin VB.Label Label10 
         Alignment       =   2  'Center
         Caption         =   "<---->"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4440
         TabIndex        =   81
         Top             =   4200
         Width           =   615
      End
      Begin VB.Label Label9 
         Alignment       =   2  'Center
         Caption         =   "<---->"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4440
         TabIndex        =   80
         Top             =   3840
         Width           =   615
      End
      Begin VB.Label Label8 
         Alignment       =   2  'Center
         Caption         =   "<---->"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4440
         TabIndex        =   79
         Top             =   3480
         Width           =   615
      End
      Begin VB.Label Label7 
         Alignment       =   2  'Center
         Caption         =   "<---->"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4440
         TabIndex        =   78
         Top             =   3120
         Width           =   615
      End
      Begin VB.Label Label6 
         Alignment       =   2  'Center
         Caption         =   "<---->"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4440
         TabIndex        =   77
         Top             =   2760
         Width           =   615
      End
      Begin VB.Label Label5 
         Alignment       =   2  'Center
         Caption         =   "<---->"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4440
         TabIndex        =   75
         Top             =   2400
         Width           =   615
      End
      Begin VB.Label Label4 
         Alignment       =   2  'Center
         Caption         =   "<---->"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4440
         TabIndex        =   74
         Top             =   2040
         Width           =   615
      End
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         Caption         =   "<---->"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4440
         TabIndex        =   73
         Top             =   1680
         Width           =   615
      End
      Begin VB.Label Label2 
         Alignment       =   2  'Center
         Caption         =   "<---->"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4440
         TabIndex        =   72
         Top             =   1320
         Width           =   615
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "<---->"
         BeginProperty Font 
            Name            =   "System"
            Size            =   9.6
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   4440
         TabIndex        =   69
         Top             =   960
         Width           =   615
      End
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear Search"
      Height          =   375
      Left            =   9840
      TabIndex        =   71
      Top             =   1320
      Width           =   1335
   End
   Begin ComctlLib.ListView lvwResults 
      Height          =   2505
      Left            =   120
      TabIndex        =   68
      Top             =   5280
      Width           =   11145
      _ExtentX        =   19664
      _ExtentY        =   4424
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      MultiSelect     =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      HideColumnHeaders=   -1  'True
      _Version        =   327682
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton cmdFind 
      Caption         =   "&Find Now"
      Default         =   -1  'True
      Height          =   375
      Left            =   9840
      TabIndex        =   70
      Top             =   720
      Width           =   1335
   End
   Begin ComctlLib.TabStrip TabStrip1 
      Height          =   5115
      Left            =   120
      TabIndex        =   1
      Top             =   45
      Width           =   9525
      _ExtentX        =   16806
      _ExtentY        =   9017
      _Version        =   327682
      BeginProperty Tabs {0713E432-850A-101B-AFC0-4210102A8DA7} 
         NumTabs         =   3
         BeginProperty Tab1 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Cross &References"
            Key             =   "Cross References"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Cross References"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab2 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "&Properties"
            Key             =   "Properties"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Properties"
            ImageVarType    =   2
         EndProperty
         BeginProperty Tab3 {0713F341-850A-101B-AFC0-4210102A8DA7} 
            Caption         =   "Results &Options"
            Key             =   "Results Options"
            Object.Tag             =   ""
            Object.ToolTipText     =   "Results Options"
            ImageVarType    =   2
         EndProperty
      EndProperty
   End
End
Attribute VB_Name = "frmQuery"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim objRecordset As New ADODB.Recordset     ' ADO Recordset used for queries
Dim objLib1 As IDMObjects.Library           ' Library 1 object
Dim objLib2 As IDMObjects.Library           ' Library 2 object
Dim objDoc As IDMObjects.Document           ' Document object
Dim colAllLibs As New IDMObjects.ObjectSet  ' Collection of all libraries in neighborhood
Dim colLib1Props As New IDMObjects.ObjectSet   ' Collection of all properties (fields) for Library 1
Dim colLib2Props As New IDMObjects.ObjectSet   ' Collection of all properties (fields) for Library 1
Dim colLib1Select As IDMObjects.ObjectSet      ' Collection of display fields for Library 1 results
Dim colLib2Select As IDMObjects.ObjectSet      ' Collection of display fields for Library 2 results
Dim colResultProps As IDMObjects.ObjectSet     ' Collection of result list properties
Dim bXRefChange As Boolean                  ' Flag for change of cross reference fields
Dim nMsgRet As Integer                      ' General message box return variable
'**********************************************************
' Procedure to select the Library 1 cross reference field *
'**********************************************************
Private Sub cboLib1XRef_Click(Index As Integer)
    
    Dim objPropDesc As IDMObjects.PropertyDescription   ' Property description
    Dim strTip As String    ' Tooltip help for field combo box
    
    ' Check that we have an item selected, other than item 0 (blank)
    If cboLib1XRef(Index).ListIndex > 0 Then
    
        ' Get the data type string for the field
        Set objPropDesc = colLib1Props.Item(cboLib1XRef(Index).ItemData(cboLib1XRef(Index).ListIndex))
        strTip = GetTypeString(objPropDesc)
            
    Else
        strTip = ""
    End If
    
    ' Set the string as the tooltip help for the field
    cboLib1XRef(Index).ToolTipText = strTip
    bXRefChange = True      ' Note that the cross refs have changed
    
End Sub
'**********************************************************
' Procedure to load the cross reference fields for        *
' Library 1                                               *
'**********************************************************
Private Sub cboLib1XRef_DropDown(Index As Integer)

    Dim nCounter As Integer     ' General counter
    Dim nFieldIndex As Integer  ' Field position reference (within all properties)
    Dim bAddIt As Boolean       ' Flag to indicate addition of field to combo required
    Dim objPropDesc As IDMObjects.PropertyDescription   ' property description
    
    Me.MousePointer = vbHourglass
    
    ' Clear the field combo box and add a blank entry as item 0
    cboLib1XRef(Index).Clear
    cboLib1XRef(Index).AddItem ""
    cboLib1XRef(Index).ItemData(cboLib1XRef(Index).NewIndex) = 0
    
    ' Iterate through the index fields (properties) for the first library
    nFieldIndex = 0
    For Each objPropDesc In colLib1Props
    
        nFieldIndex = nFieldIndex + 1
        bAddIt = True
        
        ' Check that the current field is not already selected in
        ' another combo box in the control array
        For nCounter = 0 To cboLib1XRef.Count - 1
            If cboLib1XRef(nCounter).Text = objPropDesc.Name Then
                bAddIt = False
            End If
        Next
        
        ' Check that the field is not the identity field
        If objLib1.SystemType = idmSysTypeIS And objPropDesc.Name = "F_DOCNUMBER" Then
            bAddIt = False
        ElseIf objLib1.SystemType = idmSysTypeDS And objPropDesc.Name = "idmId" Then
            bAddIt = False
        End If
            
        ' Add the field to combo box and a data field identifying the ordinal
        ' position of the field within the collection of all fields for
        ' this library (for easier ref)
        If bAddIt Then
            cboLib1XRef(Index).AddItem objPropDesc.Name
            cboLib1XRef(Index).ItemData(cboLib1XRef(Index).NewIndex) = nFieldIndex
        End If
    Next
    
    ' Select the first (blank) field, by default
    cboLib1XRef(Index).ListIndex = 0
    
    Me.MousePointer = vbDefault

End Sub
'**********************************************************
' Procedure to handle the selection of library name for   *
' Library 1                                               *
'**********************************************************
' Proc to handle the selection of library name for Lib1
Private Sub cboLib1_Click()

    Dim bFlag As Boolean        ' Logon flag
    Dim strCurrLib As String    ' String to hold library name

    On Error GoTo LogonErrorHandler
    
    ' Make sure that we have a library name selected and log on to the
    ' library
    If cboLib1.Text <> "" Then
        Set objLib1 = colAllLibs.Item(cboLib1.ItemData(cboLib1.ListIndex))
        bFlag = objLib1.Logon("", "", "", idmLogonOptWithUI)
    End If
    
    On Error GoTo PropsErrorHandler
    
    ' If both libraries are now logged on, get their index field details
    If cboLib1.Text <> "" And cboLib2.Text <> "" Then
        If objLib1.GetState(idmLibraryLoggedOn) And objLib2.GetState(idmLibraryLoggedOn) Then
            'strCurrLib = objLib1.Name
            'Set colLib1Props = objLib1.FilterPropertyDescriptions(idmObjTypeDocument, Null)
            'strCurrLib = objLib2.Name
            'Set colLib2Props = objLib2.FilterPropertyDescriptions(idmObjTypeDocument, Null)
            Dim propdescs As IDMObjects.PropertyDescriptions
            Dim propdesc As IDMObjects.PropertyDescription
            Set propdescs = objLib1.FilterPropertyDescriptions(idmObjTypeDocument, Null)
            For Each propdesc In propdescs
                colLib1Props.Add propdesc
            Next propdesc
            
            Set propdescs = objLib2.FilterPropertyDescriptions(idmObjTypeDocument, Null)
            For Each propdesc In propdescs
                colLib2Props.Add propdesc
            Next propdesc

            ' Remove any non-searchable fields
            Call RubLib1
            Call RubLib2
        End If
    End If

Exit Sub

LogonErrorHandler:

    nMsgRet = MsgBox("Error logging on to library: " & cboLib1.Text, _
                     vbExclamation + vbOKOnly)
                    
Exit Sub
   
PropsErrorHandler:

    nMsgRet = MsgBox("Error getting properties from library: " & strCurrLib, _
                    vbExclamation + vbOKOnly)
   
End Sub
'**********************************************************
' Procedure to handle the loading of available library    *
' names for Library 1                                     *
'**********************************************************
Private Sub cboLib1_DropDown()

    Dim bAddIt As Boolean       ' Flag to denote addition of library to combo
    Dim nCounter As Integer     ' General counter
    
    Me.MousePointer = vbHourglass
    
    ' Clear any current entries
    cboLib1.Clear
    
    'Iterate through all libraries
    For nCounter = 1 To colAllLibs.Count
        bAddIt = True
   
        ' Ensure that the current library is not already selected
        ' in the other combo box
        If cboLib2.Text = colAllLibs.Item(nCounter).Name Then
            bAddIt = False
        End If
                
        ' Add the library name and a data field containing the
        ' ordinal number of the library in the AllLibs collection
        ' (for easier ref)
        If bAddIt Then
            cboLib1.AddItem colAllLibs.Item(nCounter).Name
            cboLib1.ItemData(cboLib1.NewIndex) = nCounter
        End If
    Next
    
    Me.MousePointer = vbDefault

End Sub
'**********************************************************
' Procedure to load the display field list for Library 1  *
'**********************************************************
Private Sub cboLib1Select_DropDown(Index As Integer)

    Dim nCounter As Integer         ' General counter
    Dim nFieldIndex As Integer      ' Field position reference
    Dim bAddIt As Integer           ' Field position reference (within all properties)
    Dim objPropDesc As IDMObjects.PropertyDescription   ' Property description
    
    Me.MousePointer = vbHourglass
    
    ' Clear the field combo box and add a blank entry as item 0
    cboLib1Select(Index).Clear
    cboLib1Select(Index).AddItem ""
    cboLib1Select(Index).ItemData(cboLib1Select(Index).NewIndex) = 0
    
    ' Iterate through the index fields (properties) for this library
    nFieldIndex = 0
    For Each objPropDesc In colLib1Props
    
        bAddIt = True
        nFieldIndex = nFieldIndex + 1
        
        ' Check that the current field is not already selected in
        ' another combo box in the control array
        For nCounter = 0 To cboLib1Select.Count - 1
            If cboLib1Select(nCounter).Text = objPropDesc.Name Then
                bAddIt = False
            End If
        Next
        
        ' Check that the field is not the identity field
        If objLib1.SystemType = idmSysTypeIS And objPropDesc.Name = "F_DOCNUMBER" Then
            bAddIt = False
        ElseIf objLib1.SystemType = idmSysTypeDS And objPropDesc.Name = "idmId" Then
            bAddIt = False
        End If
        
        ' Add the field to combo box and a data field identifying the ordinal
        ' position of the field within the collection of all fields for
        ' this library (for easier ref)
        If bAddIt Then
            cboLib1Select(Index).AddItem objPropDesc.Name
            cboLib1Select(Index).ItemData(cboLib1Select(Index).NewIndex) = nFieldIndex
        End If
    Next
    
    ' Select the first (blank) field, by default
    cboLib1Select(Index).ListIndex = 0
    
    Me.MousePointer = vbDefault

End Sub
'**********************************************************
' Procedure to handle the selection of library name for   *
' Library 2                                               *
'**********************************************************
Private Sub cboLib2_Click()

    Dim bFlag As Boolean        ' Logon flag
    Dim strCurrLib As String    ' Library name string
    Dim nCounter As Integer     ' General counter
    
    On Error GoTo LogonErrorHandler
    
    ' Clear any current cross references
    For nCounter = 0 To cboLib1XRef.Count - 1
        If cboLib1XRef(nCounter).ListIndex > 0 Then
            cboLib1XRef(nCounter).ListIndex = 0
        End If
    Next
    For nCounter = 0 To cboLib2XRef.Count - 1
        If cboLib2XRef(nCounter).ListIndex > 0 Then
            cboLib2XRef(nCounter).ListIndex = 0
        End If
    Next
    
    ' Make sure that we have a library name selected and log on to the
    ' library
    If cboLib2.Text <> "" Then
        Set objLib2 = colAllLibs.Item(cboLib2.ItemData(cboLib2.ListIndex))
        bFlag = objLib2.Logon("", "", "", idmLogonOptWithUI)
    End If

    On Error GoTo PropsErrorHandler
    
    ' If both libraries are now logged on, get their index field details
    If cboLib1.Text <> "" And cboLib2.Text <> "" Then
        If objLib1.GetState(idmLibraryLoggedOn) And objLib2.GetState(idmLibraryLoggedOn) Then
            'strCurrLib = objLib1.Name
            'Set colLib1Props = objLib1.FilterPropertyDescriptions(idmObjTypeDocument, Null)
            Dim propdescs As IDMObjects.PropertyDescriptions
            Dim propdesc As IDMObjects.PropertyDescription
            Set propdescs = objLib1.FilterPropertyDescriptions(idmObjTypeDocument, Null)
            For Each propdesc In propdescs
                colLib1Props.Add propdesc
            Next propdesc
            
            'strCurrLib = objLib2.Name
            'Set colLib2Props = objLib2.FilterPropertyDescriptions(idmObjTypeDocument, Null)
            Set propdescs = objLib2.FilterPropertyDescriptions(idmObjTypeDocument, Null)
            For Each propdesc In propdescs
                colLib2Props.Add propdesc
            Next propdesc
            
            ' Remove any non-searchable fields
            Call RubLib1
            Call RubLib2
        End If
    End If
Exit Sub

LogonErrorHandler:

    nMsgRet = MsgBox("Error logging on to library: " & cboLib2.Text, _
                    vbExclamation + vbOKOnly)
                    
Exit Sub
   
PropsErrorHandler:

    nMsgRet = MsgBox("Error getting properties from library: " & strCurrLib, _
                    vbExclamation + vbOKOnly)

End Sub
'**********************************************************
' Procedure to handle the loading of available library    *
' names for Library 2                                     *
'**********************************************************
Private Sub cboLib2_DropDown()

    Dim bAddIt As Boolean       ' Flag to denote addition of library to combo
    Dim nCounter As Integer     ' General counter
    
    Me.MousePointer = vbHourglass
    
    ' Clear any current entries
    cboLib2.Clear
    
    'Iterate through all libraries
    For nCounter = 1 To colAllLibs.Count
        bAddIt = True
    
        ' Ensure that the current library is not already selected
        ' in the other combo box
        If cboLib1.Text = colAllLibs.Item(nCounter).Name Then
            bAddIt = False
        End If
                
        ' Add the library name and a data field containing the
        ' ordinal number of the library in the AllLibs collection
        ' (for easier ref)
        If bAddIt Then
            cboLib2.AddItem colAllLibs.Item(nCounter).Name
            cboLib2.ItemData(cboLib2.NewIndex) = nCounter
        End If
    Next
    
    Me.MousePointer = vbDefault

End Sub
'**********************************************************
' Procedure to select the Library 2 cross reference       *
'**********************************************************
Private Sub cboLib2XRef_Click(Index As Integer)

    Dim objPropDesc As IDMObjects.PropertyDescription   ' Property description
    Dim strTip As String    ' Tooltip string for field
    
    ' Check that we have an item selected, other than item 0 (blank)
    If cboLib2XRef(Index).ListIndex > 0 Then
    
        ' Get the data type string for the field
        Set objPropDesc = colLib2Props.Item(cboLib2XRef(Index).ItemData(cboLib2XRef(Index).ListIndex))
        strTip = GetTypeString(objPropDesc)
        
    Else
        strTip = ""
    End If
    
    
    ' Set the string as the tooltip help for the field
    cboLib2XRef(Index).ToolTipText = strTip
    bXRefChange = True      ' Note that the cross refs have changed

End Sub
'**********************************************************
' Procedure to load the cross reference fields for        *
' Library 2                                               *
'**********************************************************
Private Sub cboLib2XRef_DropDown(Index As Integer)

    Dim nCounter As Integer     ' General counter
    Dim nFieldIndex As Integer  ' Field position reference (within all props)
    Dim bAddIt As Boolean       ' Flag indicating whether to add field to combo
    Dim objPropDesc As IDMObjects.PropertyDescription   ' Property description
    
    Me.MousePointer = vbHourglass
    
    ' Clear the field combo box and add a blank entry as item 0
    cboLib2XRef(Index).Clear
    cboLib2XRef(Index).AddItem ""
    cboLib2XRef(Index).ItemData(cboLib2XRef(Index).NewIndex) = 0
    
    ' Iterate through the index fields (properties) for the first library
    nFieldIndex = 0
    For Each objPropDesc In colLib2Props
    
        nFieldIndex = nFieldIndex + 1
        bAddIt = True
        
        ' Check that the current field is not already selected in
        ' another combo box in the control array
        For nCounter = 0 To cboLib2XRef.Count - 1
            If cboLib2XRef(nCounter).Text = objPropDesc.Name Then
                bAddIt = False
            End If
        Next
        
        ' Check that the field is not the identity field
        If objLib2.SystemType = idmSysTypeIS And objPropDesc.Name = "F_DOCNUMBER" Then
            bAddIt = False
        ElseIf objLib2.SystemType = idmSysTypeDS And objPropDesc.Name = "idmId" Then
            bAddIt = False
        End If
        
         ' Add the field to combo box and a data field identifying the ordinal
        ' position of the field within the collection of all fields for
        ' this library (for easier ref)
       If bAddIt Then
            cboLib2XRef(Index).AddItem objPropDesc.Name
            cboLib2XRef(Index).ItemData(cboLib2XRef(Index).NewIndex) = nFieldIndex
        End If
    Next
    
    ' Select the first (blank) field, by default
    cboLib2XRef(Index).ListIndex = 0
    
    Me.MousePointer = vbDefault

End Sub
'**********************************************************
' Procedure to load the display field list for Library 2  *
'**********************************************************
Private Sub cboLib2Select_DropDown(Index As Integer)

    Dim nCounter As Integer     ' General counter
    Dim nFieldIndex As Integer  ' Field position reference (within all props)
    Dim bAddIt As Boolean       ' Flag to indicate whether to add field to combo
    Dim objPropDesc As IDMObjects.PropertyDescription   ' Property description
    
    Me.MousePointer = vbHourglass
    
    ' Clear the field combo box and add a blank entry as item 0
    cboLib2Select(Index).Clear
    cboLib2Select(Index).AddItem ""
    cboLib2Select(Index).ItemData(cboLib2Select(Index).NewIndex) = 0
    
    ' Iterate through the index fields (properties) for this library
    nFieldIndex = 0
    For Each objPropDesc In colLib2Props
    
        bAddIt = True
        nFieldIndex = nFieldIndex + 1
        
        ' Check that the current field is not already selected in
        ' another combo box in the control array
        For nCounter = 0 To cboLib2Select.Count - 1
            If cboLib2Select(nCounter).Text = objPropDesc.Name Then
                bAddIt = False
            End If
        Next
        
        ' Check that the field is not the identity field
        If objLib2.SystemType = idmSysTypeIS And objPropDesc.Name = "F_DOCNUMBER" Then
            bAddIt = False
        ElseIf objLib2.SystemType = idmSysTypeDS And objPropDesc.Name = "idmId" Then
            bAddIt = False
        End If
        
        ' Add the field to combo box and a data field identifying the ordinal
        ' position of the field within the collection of all fields for
        ' this library (for easier ref)
        If bAddIt Then
            cboLib2Select(Index).AddItem objPropDesc.Name
            cboLib2Select(Index).ItemData(cboLib2Select(Index).NewIndex) = nFieldIndex
        End If
    Next
    
    ' Select the first (blank) field, by default
    cboLib2Select(Index).ListIndex = 0
    
    Me.MousePointer = vbDefault

End Sub
'**********************************************************
' Procedure to handle the selection of a search field     *
' name                                                    *
'**********************************************************
Private Sub cboSearchField_Click(Index As Integer)
    
    Dim objPropDesc As IDMObjects.PropertyDescription   ' Property description
    Dim strTip As String    ' Tooltip string for field
    
    strTip = ""
    cboLogOp(Index).Clear
    
    ' If we have a valid search field name
    If cboSearchField(Index).Text <> "" Then
    
        Set objPropDesc = _
        colLib1Props.Item(cboSearchField(Index).ItemData(cboSearchField(Index).ListIndex))
        
        strTip = GetTypeString(objPropDesc)
       
        ' Load the corresponding logical operator combo box
        cboLogOp(Index).AddItem "="
        cboLogOp(Index).AddItem "<>"
        cboLogOp(Index).AddItem ">="
        cboLogOp(Index).AddItem "<="
        cboLogOp(Index).AddItem ">"
        cboLogOp(Index).AddItem "<"
                
        cboLogOp(Index).ListIndex = 0

    End If
    
    txtSearchData(Index).ToolTipText = strTip
    
End Sub
'**********************************************************
' Procedure to handle the loading of the query form       *
'**********************************************************
Private Sub Form_Load()

    Dim objNeighborhood As IDMObjects.Neighborhood  ' FileNet neighborhood
   
    On Error GoTo ErrorHandler
    
    ' Get an object set of all of the available libraries
    Set objNeighborhood = CreateObject("IDMObjects.Neighborhood")
    Set colAllLibs = objNeighborhood.Libraries
    
    ' Set the other object sets
    Set colLib1Select = CreateObject("IDMObjects.ObjectSet")
    Set colLib2Select = CreateObject("IDMObjects.ObjectSet")
    Set colResultProps = CreateObject("IDMObjects.ObjectSet")
    
    ' Make the first tab frame (XRef spec) visible be default
    Frame1.Visible = True
    Frame2.Visible = False
    Frame3.Visible = False
    
    bXRefChange = True      ' Force the search field lists to reload
    frmQuery.Height = 5625  ' Set the initial form height (short)

Exit Sub

ErrorHandler:

    nMsgRet = MsgBox("Error during initialization", vbCritical + vbOKOnly)
    Unload Me
    
End Sub
'**********************************************************
' Procedure to handle the loading of the query form       *
'**********************************************************
Private Sub Form_Unload(Cancel As Integer)
   
    Dim objLib As IDMObjects.Library    ' General library
    Dim nCounter As Integer             ' General counter
   
    ' Log off any libraries we've logged onto
    For Each objLib In colAllLibs
        If objLib.GetState(idmLibraryLoggedOn) Then
            objLib.Logoff
        End If
    Next
    
    ' Clear out all object sets used
    For nCounter = 1 To colAllLibs.Count
        colAllLibs.Remove 1
    Next
    For nCounter = 1 To colLib1Props.Count
        colLib1Props.Remove 1
    Next
    For nCounter = 1 To colLib2Props.Count
        colLib2Props.Remove 1
    Next
    For nCounter = 1 To colLib1Select.Count
        colLib1Select.Remove 1
    Next
    For nCounter = 1 To colLib2Select.Count
        colLib2Select.Remove 1
    Next
    For nCounter = 1 To colResultProps.Count
        colResultProps.Remove 1
    Next
    
End Sub

'**********************************************************
' Procedure to handle the selection of a results list     *
' column                                                  *
'**********************************************************
Private Sub lvwResults_ColumnClick(ByVal ColumnHeader As ComctlLib.ColumnHeader)

    ' Sort the results list according to the selected column,
    ' if the column is already selected then reverse the sort order
    If lvwResults.SortKey = ColumnHeader.Index - 1 Then
        lvwResults.SortKey = ColumnHeader.Index - 1
        lvwResults.Sorted = True
        If lvwResults.SortOrder = lvwAscending Then
            lvwResults.SortOrder = lvwDescending
        Else
            lvwResults.SortOrder = lvwAscending
        End If
    Else
        lvwResults.SortKey = ColumnHeader.Index - 1
        lvwResults.Sorted = True
        lvwResults.SortOrder = lvwAscending
    End If
    
End Sub
'**********************************************************
' Procedure to handle the selection of an item from the   *
' results list                                            *
'**********************************************************
Private Sub lvwResults_ItemClick(ByVal Item As ListItem)

    ' Create a document object form the selected item
    If Item.SubItems(1) = objLib1.Name Then
        Set objDoc = objLib1.GetObject(idmObjTypeDocument, Item)
    Else
        Set objDoc = objLib2.GetObject(idmObjTypeDocument, Item)
    End If
    
    ' Just launch the document in the IDM viewer
    objDoc.Launch
    
End Sub
'**********************************************************
' Procedure to clear the current search                   *
'**********************************************************
Private Sub cmdClear_Click()
    
    Dim nCounter As Integer     ' General counter
    
    ' Get rid of any previous items and headers
    For nCounter = 1 To lvwResults.ColumnHeaders.Count
        lvwResults.ColumnHeaders.Remove 1
    Next
    For nCounter = 1 To lvwResults.ListItems.Count
        lvwResults.ListItems.Remove 1
    Next

    ' Get rid of any search field entries
    For nCounter = 0 To cboSearchField.Count - 1
        If cboSearchField(nCounter).ListIndex > 0 Then
            cboSearchField(nCounter).ListIndex = 0
        End If
        txtSearchData(nCounter).Text = ""
    Next
    
    'Get rid of any select entries
    For nCounter = 0 To cboLib1Select.Count - 1
        If cboLib1Select(nCounter).ListIndex > 0 Then
            cboLib1Select(nCounter).ListIndex = 0
        End If
    Next
    For nCounter = 0 To cboLib2Select.Count - 1
        If cboLib2Select(nCounter).ListIndex > 0 Then
            cboLib2Select(nCounter).ListIndex = 0
        End If
    Next
    
    ' Resize the query form to get rid of the
    ' results list
    frmQuery.Height = 5625
    
End Sub
'**********************************************************
' Procedure to handle the switching of the relational     *
' operator                                                *
'**********************************************************
Private Sub cmdRelOp_Click(Index As Integer)
    
    ' Switch the relational operator for the search
    If cmdRelOp(Index).Caption = "AND" Then
        cmdRelOp(Index).Caption = "OR"
    Else
        cmdRelOp(Index).Caption = "AND"
    End If
    
End Sub
'**********************************************************
' Procedure to invoke the search                          *
'**********************************************************
Private Sub cmdFind_Click()
    
    ' If we have valid search conditions, set up and run
    ' the search against both libraries
    If ValidateSearch Then
        Call GetLib1Selects
        Call GetLib2Selects
        Call GetResultsProps
        Call SetupLibResultsList
        Call SearchLib1
        Call SearchLib2
        
        ' Set the form height to accomodate the results list
        frmQuery.Height = 8325
    End If

End Sub
'**********************************************************
' Procedure to handle the selection of a tab              *
'**********************************************************
Private Sub TabStrip1_Click()
    
    ' See which tab has been selected
    Select Case TabStrip1.SelectedItem.Key
    
        Case "Cross References"
            Frame1.Visible = True
            Frame2.Visible = False
            Frame3.Visible = False
            
        Case "Properties"
            ' See if any of the cross refs have changed, if so,
            ' we need to reload the avaliable search field list
            If bXRefChange Then
                Call LoadFieldList
                bXRefChange = False
            End If
            Frame2.Visible = True
            Frame1.Visible = False
            Frame3.Visible = False
            
         Case "Results Options"
            lblLib1.Caption = cboLib1.Text
            lblLib2.Caption = cboLib2.Text
            Frame3.Visible = True
            Frame1.Visible = False
            Frame2.Visible = False
           
        Case Else
            Frame2.Visible = False
            Frame1.Visible = False
            Frame3.Visible = False
            
    End Select
            
End Sub
'**********************************************************
' Procedure to load the available search fields           *
'**********************************************************
Private Sub LoadFieldList()

    Dim nXRefIndex As Integer       ' Cross reference field position (in all props)
    Dim nSearchIndex As Integer     ' Search field counter
    Dim objPropDesc As IDMObjects.PropertyDescription   ' Property description
    Dim nCounter As Integer         ' General counter
    
    ' Clear all of the search fields and add a blank entry as the default
    For nSearchIndex = 0 To cboSearchField.Count - 1
        cboSearchField(nSearchIndex).Clear
        cboSearchField(nSearchIndex).AddItem ""
        cboSearchField(nSearchIndex).ItemData(cboSearchField(nSearchIndex).NewIndex) = 0
    Next
    
    ' Iterate through each of the cross references
    For nXRefIndex = 0 To cboLib1XRef.Count - 1
    
        ' Ensure that the cross ref is valid
        If cboLib1XRef(nXRefIndex).Text <> "" And cboLib2XRef(nXRefIndex) <> "" Then
        
            ' Add the cross ref to each of the search fields
            For nSearchIndex = 0 To cboSearchField.Count - 1
                
                ' If we have the first (Key) search field, only add to this
                ' cross references which are key fields in both libraries
                If nSearchIndex = 0 Then
                    Set objPropDesc = colLib1Props.Item(cboLib1XRef(nXRefIndex).ItemData(cboLib1XRef(nXRefIndex).ListIndex))
                    If objPropDesc.GetState(idmPropKey) Then
                        Set objPropDesc = colLib2Props.Item(GetLib2XRef(cboLib1XRef(nXRefIndex).Text))
                        If objPropDesc.GetState(idmPropKey) Then
                            cboSearchField(nSearchIndex).AddItem cboLib1XRef(nXRefIndex)
                            cboSearchField(nSearchIndex).ItemData(cboSearchField(nSearchIndex).NewIndex) = cboLib1XRef(nXRefIndex).ItemData(cboLib1XRef(nXRefIndex).ListIndex)
                        End If
                    End If
                Else
                    cboSearchField(nSearchIndex).AddItem cboLib1XRef(nXRefIndex)
                    cboSearchField(nSearchIndex).ItemData(cboSearchField(nSearchIndex).NewIndex) = cboLib1XRef(nXRefIndex).ItemData(cboLib1XRef(nXRefIndex).ListIndex)
                End If
            Next
            
        End If
        
    Next
    
    ' Select the default field
    For nSearchIndex = 0 To cboSearchField.Count - 1
        cboSearchField(nSearchIndex).ListIndex = 0
        cboLogOp(nSearchIndex).Clear
        txtSearchData(nSearchIndex).Text = ""
    Next
    
End Sub
'**********************************************************
' Procedure to validate the search specification          *
'**********************************************************
Private Function ValidateSearch() As Boolean

    Dim nCounter As Integer     ' General counter
    Dim bValid As Boolean       ' Validity flag
    
    bValid = True
    
    ' Check that we are logged on to both libraries
    If cboLib1.Text = "" Or cboLib2.Text = "" Then
        nMsgRet = MsgBox("You must select two libraries before running a search", _
                    vbExclamation + vbOKOnly, "Search Validation Error")
        bValid = False
    End If
    
    ' Check through each of the search conditions
    For nCounter = 0 To cboSearchField.Count - 1
        ' See if we have a field name with no data entered
        If cboSearchField(nCounter).Text <> "" And _
          Trim(txtSearchData(nCounter).Text = "") Then
            nMsgRet = MsgBox("Data field for " & cboSearchField(nCounter).Text & _
                        " can not be empty", vbExclamation + vbOKOnly, _
                        "Search Validation Error")
            txtSearchData(nCounter).SetFocus
            bValid = False
        End If
    Next
    
    ValidateSearch = bValid
    
End Function
'**********************************************************
' Procedure to remove any non-searchable or non-          *
' selectable fields from the available fields collection  *
' for library 2 (only applies to Mezzanine fields         *
'**********************************************************
Private Sub RubLib2()

    Dim nCounter As Integer     ' General counter
    
    ' Get rid of non-searchable Mezzanine fields from the index field list
    If objLib2.SystemType = idmSysTypeDS Then
    
        Me.MousePointer = vbHourglass
        
        Dim propdesc As IDMObjects.PropertyDescription
        For nCounter = colLib2Props.Count To 1 Step -1
            Set propdesc = colLib2Props.Item(nCounter)
  'DEBUG BEGIN
  '          If propdesc.Name = "idmDocType" Then
  '              Dim b As Boolean
  '              b = propdesc.GetState(idmPropSearchable)
  '              b = propdesc.GetState(idmPropSelectable)
  '              MsgBox (propdesc.Name)
  '          End If
  'DEBUG END
  'GCL - GetExtendedProperty to GetState
            'If (Not colLib2Props.Item(nCounter).GetExtendedProperty("F_ISSELECTABLE")) Or _
            '    (Not colLib2Props.Item(nCounter).GetExtendedProperty("F_ISSEARCHABLE")) Then
            If (Not propdesc.GetState(idmPropSelectable)) Or _
                (Not propdesc.GetState(idmPropSearchable)) Then
                colLib2Props.Remove nCounter
            End If
        Next
       
        Me.MousePointer = vbDefault
        
    End If

End Sub
'**********************************************************
' Procedure to return the ordinal position of a field in  *
' the collection of library 2 fields, based on the name   *
' of a field in library 1                                 *
'**********************************************************
Private Function GetLib2XRef(strLib1Field As String) As Integer

    Dim nCounter As Integer     ' General counter
    Dim nLib2Index As Integer   ' Field position (in library 2)
    Dim bLoop As Boolean        ' Loop flag
    
    bLoop = True
    nCounter = 0
    nLib2Index = 0
    
    ' Ensure that we've passed in a valid field name
    If Not IsNull(strLib1Field) Then
        If strLib1Field = "" Then
            bLoop = False
        End If
    Else
        bLoop = False
    End If
    
    ' Loop through searching for the cross ref (if not found,
    ' we will return zero)
    While bLoop
        If cboLib1XRef(nCounter).Text = strLib1Field Then
            nLib2Index = cboLib2XRef(nCounter).ItemData(cboLib2XRef(nCounter).ListIndex)
            bLoop = False
        Else
            nCounter = nCounter + 1
        End If
        
        If nCounter = cboLib2XRef.Count Then
            bLoop = False
        End If
    Wend
    
    GetLib2XRef = nLib2Index
    
End Function
'**********************************************************
' Procedure to build a select string (to be used as a SQL *
' string component) for library 2                         *
'**********************************************************
Private Function BuildLib2Select() As String

    Dim strSelect As String     ' SQL select string portion
    Dim nCounter As Integer     ' General counter
    Dim objPropDesc As IDMObjects.PropertyDescription   ' Property description
    
    ' Add the appropriate identity field name
    If objLib2.SystemType = idmSysTypeIS Then
        strSelect = " F_DOCNUMBER"
    Else
        strSelect = " idmId"
    End If
    
    ' Iterate through each field in the library 2
    ' selection collection
    For Each objPropDesc In colLib2Select
        strSelect = strSelect & ", " & objPropDesc.Name
    Next
    
    BuildLib2Select = strSelect

End Function
'**********************************************************
' Procedure to build a select string (to be used as a SQL *
' string component) for library 1                         *
'**********************************************************
Private Function BuildLib1Select() As String

    Dim strSelect As String     ' SQL Select string portion
    Dim nCounter As Integer     ' General counter
    Dim objPropDesc As IDMObjects.PropertyDescription   ' Property description
    
    ' Add the appropriate identity field name
    If objLib1.SystemType = idmSysTypeIS Then
        strSelect = " F_DOCNUMBER"
    Else
        strSelect = " idmId"
    End If
    
    ' Iterate through each field in the library 1
    ' selection collection
    For Each objPropDesc In colLib1Select
        strSelect = strSelect & ", " & objPropDesc.Name
    Next
    
    BuildLib1Select = strSelect

End Function
'**********************************************************
' Procedure to build a where string (to be used as a SQL  *
' string component) for library 1                         *
'**********************************************************
Private Function BuildLib1Filter() As String

    Dim strFilter As String     ' SQL string portion
    Dim nCounter As Integer     ' General counter
    Dim objPropDesc As IDMObjects.PropertyDescription   ' Property description
    Dim bFilter As Boolean      ' Filter flag (we have some filter conditions)
    
    ' See whether we have to query an IMS library
    If objLib1.SystemType = idmSysTypeIS Then
    
        bFilter = False
        strFilter = ""
        
        ' Iterate through each search field on the query tab
        For nCounter = 0 To cboSearchField.Count - 1
            If cboSearchField(nCounter).Text <> "" Then
                If nCounter = 1 Then
                    bFilter = True
                    ' See if we have a key specified, if so, this
                    ' will always be ANDed to other search conditions
                    If cboSearchField(0).Text <> "" Then
                        strFilter = strFilter & " AND ("
                    End If
                End If
                
                ' Add the relational operator
                If nCounter > 1 Then
                    strFilter = strFilter & " " & cmdRelOp(nCounter - 2).Caption
                End If
                
                ' Add the field name and logical operator
                Set objPropDesc = colLib1Props.Item(cboSearchField(nCounter).ItemData(cboSearchField(nCounter).ListIndex))
                strFilter = strFilter & " " & objPropDesc.Name & " " & _
                            cboLogOp(nCounter).Text & " "
                
                ' Add the search criterion string, according to type
                Select Case objPropDesc.TypeID
                    Case idmTypeString
                        strFilter = strFilter & "'" & txtSearchData(nCounter).Text & "'"
                        
                    Case idmTypeDate
                        strFilter = strFilter & "'" & txtSearchData(nCounter).Text & "'"
                    
                    Case Else
                        strFilter = strFilter & txtSearchData(nCounter).Text
                End Select
            End If
        Next
        
        ' Delimit if we have a key + filter situation
        If cboSearchField(0).Text <> "" And bFilter Then
            strFilter = strFilter & " )"
        End If
    
    Else
        strFilter = ""
        
        ' Iterate through each search field on the query tab
        For nCounter = 0 To cboSearchField.Count - 1
            If cboSearchField(nCounter).Text <> "" Then
                If nCounter = 1 Then
                    ' See if we have a key specified, if so, this
                    ' will always be ANDed to other search conditions
                   If cboSearchField(0).Text <> "" Then
                        strFilter = strFilter & " AND ("
                    End If
                End If
                
                ' Add the relational operator
                If nCounter > 1 Then
                    strFilter = strFilter & " " & cmdRelOp(nCounter - 2).Caption
                End If
                
                ' Add the field name and logical operator
                Set objPropDesc = colLib1Props.Item(cboSearchField(nCounter).ItemData(cboSearchField(nCounter).ListIndex))
                strFilter = strFilter & " " & objPropDesc.Name & " " & _
                            cboLogOp(nCounter).Text & " "
                            
                ' Add the search criterion string, according to type
                Select Case objPropDesc.TypeID
                    Case idmTypeString
                        strFilter = strFilter & "'" & txtSearchData(nCounter).Text & "'"
                        
                    Case idmTypeDate
                        strFilter = strFilter & "'" & Format(txtSearchData(nCounter).Text, "yyyy-mm-dd") & "T00:00:00,000'"
                    
                    Case Else
                        strFilter = strFilter & txtSearchData(nCounter).Text
                End Select
            End If
        Next
        
        ' Delimit if we have a key + filter situation
        If cboSearchField(0).Text <> "" Then
            strFilter = strFilter & " )"
        End If
    End If
    
    ' Add the WHERE clause, if we have any search fields
    If strFilter <> "" Then
        strFilter = " WHERE" & strFilter
    End If
    
    BuildLib1Filter = strFilter
    
End Function
'**********************************************************
' Procedure to remove any non-searchable or non-          *
' selectable fields from the available fields collection  *
' for library 1 (only applies to Mezzanine fields         *
'**********************************************************
Private Sub RubLib1()

    Dim nCounter As Integer     ' General counter
    
    ' Get rid of non-searchable Mezzanine fields from the index field list
    If objLib1.SystemType = idmSysTypeDS Then
    
        Me.MousePointer = vbHourglass
        
        Dim propdesc As IDMObjects.PropertyDescription
        For nCounter = colLib1Props.Count To 1 Step -1
            Set propdesc = colLib1Props.Item(nCounter)
  'DEBUG BEGIN
  '          If propdesc.Name = "idmDocType" Then
  '              Dim b As Boolean
  '              b = propdesc.GetState(idmPropSearchable)
  '              b = propdesc.GetState(idmPropSelectable)
  '              MsgBox (propdesc.Name)
  '          End If
  'DEBUG END
  'GCL - GetExtendedProperty to GetState
            'If (Not colLib1Props.Item(nCounter).GetExtendedProperty("F_ISSELECTABLE")) Or _
            '    (Not colLib1Props.Item(nCounter).GetExtendedProperty("F_ISSEARCHABLE")) Then
            If (Not propdesc.GetState(idmPropSelectable)) Or _
                (Not propdesc.GetState(idmPropSearchable)) Then
                colLib1Props.Remove nCounter
            End If
        Next
        
        Me.MousePointer = vbDefault
        
    End If

End Sub
'**********************************************************
' Procedure to build a where string (to be used as a SQL  *
' string component) for library 2                         *
'**********************************************************
Private Function BuildLib2Filter() As String

    Dim strFilter As String     ' SQL string portion
    Dim nCounter As Integer     ' General counter
    Dim objPropDesc As IDMObjects.PropertyDescription   ' Property description
    Dim bFilter As Boolean      ' Filter flag (we have some filter conditions)
    
    ' See whether we have to query an IMS library
    If objLib2.SystemType = idmSysTypeIS Then
    
        bFilter = False
        strFilter = ""
        
        ' Iterate through each search field on the query tab
        For nCounter = 0 To cboSearchField.Count - 1
            If cboSearchField(nCounter).Text <> "" Then
                If nCounter = 1 Then
                    bFilter = True
                    ' See if we have a key specified, if so, this
                    ' will always be ANDed to other search conditions
                    If cboSearchField(0).Text <> "" Then
                        strFilter = strFilter & " AND ("
                    End If
                End If
                
                ' Add the relational operator
                If nCounter > 1 Then
                    strFilter = strFilter & " " & cmdRelOp(nCounter - 2).Caption
                End If
                
                ' Add the field name and logical operator (the field details
                ' must be obtained from the library 1 cross ref
                Set objPropDesc = colLib2Props.Item(GetLib2XRef(cboSearchField(nCounter).Text))
                strFilter = strFilter & " " & objPropDesc.Name & " " & _
                            cboLogOp(nCounter).Text & " "
                
                ' Add the search criterion string, according to type
                Select Case objPropDesc.TypeID
                    Case idmTypeString
                        strFilter = strFilter & "'" & txtSearchData(nCounter).Text & "'"
                        
                    Case idmTypeDate
                        strFilter = strFilter & "'" & txtSearchData(nCounter).Text & "'"
                    
                    Case Else
                        strFilter = strFilter & txtSearchData(nCounter).Text
                End Select
            End If
        Next
        
        ' Delimit if we have a key + filter situation
        If cboSearchField(0).Text <> "" And bFilter Then
            strFilter = strFilter & " )"
        End If
    
    
    Else
        strFilter = ""
        
        ' Iterate through each search field on the query tab
        For nCounter = 0 To cboSearchField.Count - 1
            If cboSearchField(nCounter).Text <> "" Then
                If nCounter = 1 Then
                    ' See if we have a key specified, if so, this
                    ' will always be ANDed to other search conditions
                    If cboSearchField(0).Text <> "" Then
                        strFilter = strFilter & " AND ("
                    End If
                End If
                
                ' Add the relational operator
                If nCounter > 1 Then
                    strFilter = strFilter & " " & cmdRelOp(nCounter - 2).Caption
                End If
                
                ' Add the field name and logical operator (the field details
                ' must be obtained from the library 1 cross ref
                Set objPropDesc = colLib2Props.Item(GetLib2XRef(cboSearchField(nCounter).Text))
                strFilter = strFilter & " " & objPropDesc.Name & " " & _
                            cboLogOp(nCounter).Text & " "
                            
                ' Add the search criterion string, according to type
                Select Case objPropDesc.TypeID
                    Case idmTypeString
                        strFilter = strFilter & "'" & txtSearchData(nCounter).Text & "'"
                        
                    Case idmTypeDate
                        strFilter = strFilter & "'" & Format(txtSearchData(nCounter).Text, "yyyy-mm-dd") & "T00:00:00,000'"
                    
                    Case Else
                        strFilter = strFilter & txtSearchData(nCounter).Text
                End Select
            End If
        Next
        
        ' Delimit if we have a key + filter situation
        If cboSearchField(0).Text <> "" Then
            strFilter = strFilter & " )"
        End If
    
    End If
    
    ' Add the WHERE clause, if we have any search fields
    If strFilter <> "" Then
        strFilter = " WHERE" & strFilter
    End If

    BuildLib2Filter = strFilter
    
End Function
'**********************************************************
' Procedure to build load a search result item into the   *
' result list for library 1                               *
'**********************************************************
Private Sub LoadLib1ResultItem()

    Dim itmDetail As ListItem   ' List item object
    Dim nItemCount As Integer   ' Sub item counter
    Dim nCounter As Integer     ' General counter
    Dim objPropDesc As IDMObjects.PropertyDescription       ' Property description
    Dim objPropTestDesc As IDMObjects.PropertyDescription   ' Temp property description
    
    Set itmDetail = lvwResults.ListItems.Add()
    
    ' Add the identity field contents
    If objLib1.SystemType = idmSysTypeIS Then
        itmDetail.Text = objRecordset.Fields.Item("F_DOCNUMBER")
    Else
        itmDetail.Text = objRecordset.Fields.Item("idmId")
    End If
    
    itmDetail.SubItems(1) = objLib1.Name
    nItemCount = 2
    
    ' Iterate through the result props list
    For Each objPropDesc In colResultProps
        ' Look through each of the selection fields for the library
        For nCounter = 1 To colLib1Select.Count
            Set objPropTestDesc = colLib1Select.Item(nCounter)
            ' See if the selction field name matches the results field name
            If objPropTestDesc.Name = objPropDesc.Name Then
                ' Output the contents of this field as the next subitem
 'DEBUG BEGIN
 'Dim fld As ADODB.Field
 'MsgBox (objPropDesc.Name)
 'For Each fld In objRecordset.Fields
    'MsgBox (fld.Name)
 'Next fld
 'DEBUG END
                If Not IsNull(objRecordset.Fields.Item(objPropDesc.Name)) Then
                    itmDetail.SubItems(nItemCount) = _
                            objRecordset.Fields.Item(objPropDesc.Name)
                Else
                    itmDetail.SubItems(nItemCount) = ""
                End If
            End If
        Next
        nItemCount = nItemCount + 1
    Next
    
End Sub
'**********************************************************
' Procedure to add the result column headers (field names)*
' to the result list                                      *
'**********************************************************
Private Sub SetupLibResultsList()

    Dim nCounter As Integer         ' General counter
    Dim clmHead As ColumnHeader     ' Column header object
    Dim bAddIt As Boolean           ' Flag to determine whether to add subitem
    Dim objPropDesc As IDMObjects.PropertyDescription   ' Property description

    ' Get rid of any previous items and headers
    For nCounter = 1 To lvwResults.ColumnHeaders.Count
        lvwResults.ColumnHeaders.Remove 1
    Next
    For nCounter = 1 To lvwResults.ListItems.Count
        lvwResults.ListItems.Remove 1
    Next
    
    ' Add the identity and library columns
    Set clmHead = lvwResults.ColumnHeaders.Add()
    clmHead.Text = "Document ID"
    Set clmHead = lvwResults.ColumnHeaders.Add()
    clmHead.Text = "Library Name"
    
    ' Add a column for each field in the result fields
    ' collection
    For Each objPropDesc In colResultProps
        Set clmHead = lvwResults.ColumnHeaders.Add()
        clmHead.Text = objPropDesc.Name
    Next
    
    lvwResults.HideColumnHeaders = False

End Sub
'**********************************************************
' Procedure to return the ordinal position of a field in  *
' the collection of library 1 fields, based on the name   *
' of a field in library 2                                 *
'**********************************************************
Private Function GetLib1XRef(strLib2Field As String) As Integer

    Dim nCounter As Integer     ' General counter
    Dim nLib1Index As Integer   ' Field position reference (in library 2)
    Dim bLoop As Boolean        ' Loop flag
    
    bLoop = True
    nCounter = 0
    nLib1Index = 0
    
    ' Ensure that we've passed in a valid field name
    If Not IsNull(strLib2Field) Then
        If strLib2Field = "" Then
            bLoop = False
        End If
    Else
        bLoop = False
    End If
    
   ' Loop through searching for the cross ref (if not found,
   ' we will return zero)
   While bLoop
        If cboLib2XRef(nCounter).Text = strLib2Field Then
            nLib1Index = cboLib1XRef(nCounter).ItemData(cboLib1XRef(nCounter).ListIndex)
            bLoop = False
        Else
            nCounter = nCounter + 1
        End If
        
        If nCounter = cboLib1XRef.Count Then
            bLoop = False
        End If
    Wend
    
    GetLib1XRef = nLib1Index
    
End Function
'**********************************************************
' Procedure to create a collection of select fields from  *
' those fields selected by the user, for library 1        *
'**********************************************************
Private Sub GetLib1Selects()

    Dim nCounter As Integer     ' General counter
    Dim objPropDesc As IDMObjects.PropertyDescription   ' Property description
    Dim bAddIt As Boolean       ' Flag to denote whether to add select field
    Dim objPropTestDesc As IDMObjects.PropertyDescription   ' Temp property description
    
    ' Empty the Select field collection for this library
    For nCounter = 1 To colLib1Select.Count
        colLib1Select.Remove 1
    Next
    
    ' Iterate through each Select field, chosen by the user from Lib1
    For nCounter = 0 To cboLib1Select.Count - 1
        If cboLib1Select(nCounter).Text <> "" Then
            Set objPropDesc = colLib1Props.Item(cboLib1Select(nCounter).ItemData(cboLib1Select(nCounter).ListIndex))
            colLib1Select.Add objPropDesc
        End If
    Next
    
    ' Check the library 2 Select fields for any library 1 cross refs
    For nCounter = 0 To cboLib2Select.Count - 1
        If cboLib2Select(nCounter).Text <> "" Then
        
            If GetLib1XRef(cboLib2Select(nCounter).Text) > 0 Then
                Set objPropTestDesc = colLib1Props.Item(GetLib1XRef(cboLib2Select(nCounter).Text))
                
                ' Ensure we've not already got this in library 1
                bAddIt = True
                For Each objPropDesc In colLib1Select
                    If objPropDesc.Name = objPropTestDesc.Name Then
                        bAddIt = False
                    End If
                Next
                
                ' Add the field props to the collection
                If bAddIt Then
                    colLib1Select.Add objPropTestDesc
                End If
            End If
        End If
    Next
        
End Sub
'**********************************************************
' Procedure to create a collection of select fields from  *
' those fields selected by the user, for library 2        *
'**********************************************************
Private Sub GetLib2Selects()

    Dim nCounter As Integer     ' General counter
    Dim objPropDesc As IDMObjects.PropertyDescription   ' Property description
    Dim bAddIt As Boolean       ' Flag to denote whether select field should be added
    Dim objPropTestDesc As IDMObjects.PropertyDescription   ' Temp property description
    
    ' Empty the Select field collection for this library
    For nCounter = 1 To colLib2Select.Count
        colLib2Select.Remove 1
    Next
    
    ' Iterate through each Select field, chosen by the user from Lib2
    For nCounter = 0 To cboLib2Select.Count - 1
        If cboLib2Select(nCounter).Text <> "" Then
            Set objPropDesc = colLib2Props.Item(cboLib2Select(nCounter).ItemData(cboLib2Select(nCounter).ListIndex))
            colLib2Select.Add objPropDesc
        End If
    Next
    
    ' Check the library 1 Select fields for any library 2 cross refs
    For nCounter = 0 To cboLib1Select.Count - 1
        If cboLib1Select(nCounter).Text <> "" Then
        
            If GetLib2XRef(cboLib1Select(nCounter).Text) > 0 Then
                Set objPropTestDesc = colLib2Props.Item(GetLib2XRef(cboLib1Select(nCounter).Text))
                
                ' Ensure we've not already got this in library 2
                bAddIt = True
                For Each objPropDesc In colLib2Select
                    If objPropDesc.Name = objPropTestDesc.Name Then
                        bAddIt = False
                    End If
                Next
                
                ' Add the field props to the collection
                If bAddIt Then
                    colLib2Select.Add objPropTestDesc
                End If
            End If
        End If
    Next
        
End Sub
'**********************************************************
' Procedure to create a collection of required results    *
' fields                                                  *
'**********************************************************
Private Sub GetResultsProps()

    Dim nCounter As Integer     ' General counter
    Dim objPropDesc As IDMObjects.PropertyDescription   ' Property description
    Dim bAddIt As Boolean       ' Flag to denote whether to add field to collection
    Dim objPropTestDesc As IDMObjects.PropertyDescription   ' Temp property description
    
    ' Empty the results props list
    For nCounter = 1 To colResultProps.Count
        colResultProps.Remove 1
    Next
    
    ' Add all of the field props from the library 1 Select collection
    For Each objPropDesc In colLib1Select
        colResultProps.Add objPropDesc
    Next
        
    ' Add any Lib2 Select fields not already cross referenced
    For Each objPropDesc In colLib2Select
        bAddIt = True
        If GetLib1XRef(objPropDesc.Name) > 0 Then
            ' Ensure we've not already got this in results set
            bAddIt = True
            For Each objPropTestDesc In colResultProps
                If objPropTestDesc.Name = colLib1Props.Item(GetLib1XRef(objPropDesc.Name)).Name Then
                    bAddIt = False
                End If
            Next
            
        End If
        
        If bAddIt Then
            colResultProps.Add objPropDesc
        End If
    
    Next
        
End Sub
'**********************************************************
' Procedure to search library 1                           *
'**********************************************************
Private Sub SearchLib1()

    Dim strConnect As String    ' ADO connect string
    Dim strQuery As String      ' ADO query string
    
    On Error GoTo ErrorHandler
    
    Me.MousePointer = vbHourglass
    
    If objLib1.SystemType = idmSysTypeIS Then
     
        ' Search the IMS Library
         strConnect = "provider=FnDBProvider;data source=" _
                         & objLib1.Name _
                         & ";LogonID=" & objLib1.LogonId _
                         & ";Prompt=4;SystemType=" & objLib1.SystemType & ";"
     
         strQuery = "SELECT" & BuildLib1Select & _
                    " FROM FnDocument" & _
                    BuildLib1Filter
                    
    Else
    
        ' Search the Mezzanine Library
         strConnect = "provider=FnDBProvider;data source=" _
                         & objLib1.Name _
                         & ";User ID=;Password=" _
                         & ";Prompt=1;SystemType=" & objLib1.SystemType & ";"
         
         'WON'T WORK strQuery = "SELECT *" & '_
         'DEBUG strQuery = "SELECT idmId, idmDocType" & _
         '           " FROM FnDocument" & _
         '           BuildLib1Filter

         strQuery = "SELECT " & BuildLib1Select & _
                    " FROM FnDocument" & _
                    BuildLib1Filter
    End If
        
    ' Define a new recordset
    Set objRecordset = New ADODB.Recordset
                
    objRecordset.MaxRecords = 100
    objRecordset.Open strQuery, strConnect, adOpenKeyset
    
    ' Add each recordset row to the result list
    If objRecordset.RecordCount > 0 Then
        objRecordset.MoveFirst
        While Not objRecordset.EOF
            Call LoadLib1ResultItem
            objRecordset.MoveNext
        Wend
        
    End If
    
    objRecordset.Close
         
    Me.MousePointer = vbDefault
    
Exit Sub

ErrorHandler:

    Me.MousePointer = vbDefault
    nMsgRet = MsgBox("Error searching library : " & objLib1.Name, _
                        vbExclamation + vbOKOnly)
    
End Sub
'**********************************************************
' Procedure to search library 2                           *
'**********************************************************
Private Sub SearchLib2()

    Dim strConnect As String    ' ADO connect string
    Dim strQuery As String      ' ADO query string
    
    On Error GoTo ErrorHandler
    
    Me.MousePointer = vbHourglass
    
    If objLib2.SystemType = idmSysTypeIS Then
     
        ' Search the IMS Library
         strConnect = "provider=FnDBProvider;data source=" _
                         & objLib2.Name _
                         & ";LogonID=" & objLib2.LogonId _
                         & ";Prompt=4;SystemType=" & objLib2.SystemType & ";"
     
         strQuery = "SELECT" & BuildLib2Select & _
                    " FROM FnDocument" & _
                    BuildLib2Filter
                    
    Else
    
        ' Search the Mezzanine Library
         strConnect = "provider=FnDBProvider;data source=" _
                         & objLib2.Name _
                         & ";User ID=;Password=" _
                         & ";Prompt=1;SystemType=" & objLib2.SystemType & ";"
         
         'strQuery = "SELECT *" & _
                    " FROM FnDocument" & _
                    BuildLib2Filter
     
         strQuery = "SELECT" & BuildLib2Select & _
                    " FROM FnDocument" & _
                    BuildLib2Filter
    End If
        
    ' Define a new recordset
    Set objRecordset = New ADODB.Recordset
                
    objRecordset.MaxRecords = 100
    objRecordset.Open strQuery, strConnect, adOpenKeyset
    
    ' Add each recordset row to the result list
    If objRecordset.RecordCount > 0 Then
        objRecordset.MoveFirst
        While Not objRecordset.EOF
            Call LoadLib2ResultItem
            objRecordset.MoveNext
        Wend
        
    End If
    
    objRecordset.Close
         
    Me.MousePointer = vbDefault

Exit Sub

ErrorHandler:

    Me.MousePointer = vbDefault
    nMsgRet = MsgBox("Error searching library : " & objLib2.Name, _
                        vbExclamation + vbOKOnly)
    
End Sub
'**********************************************************
' Procedure to build load a search result item into the   *
' result list for library 2                               *
'**********************************************************
Private Sub LoadLib2ResultItem()

    Dim itmDetail As ListItem   ' List item object
    Dim nItemCount As Integer   ' Sub item counter
    Dim nCounter As Integer     ' General counter
    Dim objPropDesc As IDMObjects.PropertyDescription       ' Property description
    Dim objPropTestDesc As IDMObjects.PropertyDescription   ' Temp property description


    Set itmDetail = lvwResults.ListItems.Add()
    
    ' Add the identity field contents
    If objLib2.SystemType = idmSysTypeIS Then
        itmDetail.Text = objRecordset.Fields.Item("F_DOCNUMBER")
    Else
        itmDetail.Text = objRecordset.Fields.Item("idmId")
    End If
    
    itmDetail.SubItems(1) = objLib2.Name
    nItemCount = 2
    
    ' Iterate through the result props list
    For Each objPropDesc In colResultProps
        ' Look through each of the selection fields for the library
        For nCounter = 1 To colLib2Select.Count
            Set objPropTestDesc = colLib2Select.Item(nCounter)
            ' See if the selction field name matches the results field name
            If objPropTestDesc.Name = objPropDesc.Name Then
                ' Output the contents of this field as the next subitem
                If Not IsNull(objRecordset.Fields.Item(objPropDesc.Name)) Then
                    itmDetail.SubItems(nItemCount) = _
                            objRecordset.Fields.Item(objPropDesc.Name)
                Else
                    itmDetail.SubItems(nItemCount) = ""
                End If
            ' Otherwise, se we have a cross ref to this field
            ElseIf GetLib2XRef(objPropDesc.Name) > 0 Then
                Set objPropTestDesc = colLib2Props.Item(GetLib2XRef(objPropDesc.Name))
                ' Output the contents of this field as the next subitem
'DEBUG BEGIN
'               Dim fld As ADODB.Field
'               For Each fld In objRecordset.Fields
'                    MsgBox fld.Name
'                Next fld
'DEBUG END
                If Not IsNull(objRecordset.Fields.Item(objPropTestDesc.Name)) Then
                    itmDetail.SubItems(nItemCount) = _
                            objRecordset.Fields.Item(objPropTestDesc.Name)
                Else
                    itmDetail.SubItems(nItemCount) = ""
                End If
            End If
        Next
        
        nItemCount = nItemCount + 1
            
    Next
    
End Sub
'**********************************************************
' Procedure to return a string describing a property type *
'**********************************************************
Private Function GetTypeString(objPropDesc As IDMObjects.PropertyDescription) As String

    Dim strRet As String    ' Return string
    
    ' Produce the field type description string
    Select Case objPropDesc.TypeID
        Case idmTypeString
            strRet = "String,  Length: " & CStr(objPropDesc.Size)
            
        Case idmTypeDate
            strRet = "Date"
            
        Case idmTypeDouble, idmTypeSingle, idmTypeShort, idmTypeLong
            strRet = "Number"
            
        Case Else
            strRet = CStr(objPropDesc.TypeID)
    End Select
    strRet = "Type: " & strRet
    
    ' Add details of whether the property is a key field or not
    If objPropDesc.GetState(idmPropKey) Then
        strRet = strRet & ", Key: True"
    Else
        strRet = strRet & ", Key: False"
    End If
    
    ' Return the completed string
    GetTypeString = strRet
    
End Function

