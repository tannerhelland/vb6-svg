VERSION 5.00
Begin VB.Form frmDemo 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   Caption         =   "SVG support in VB6 - https://github.com/tannerhelland/vb6-svg"
   ClientHeight    =   8325
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   14745
   BeginProperty Font 
      Name            =   "Segoe UI"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   555
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   983
   StartUpPosition =   3  'Windows Default
   Begin VB.CheckBox chkFit 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "auto-fit to picture box"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   240
      TabIndex        =   13
      Top             =   4800
      Value           =   1  'Checked
      Width           =   3255
   End
   Begin VB.CheckBox chkAutoClear 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "clear picture box "
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   240
      TabIndex        =   12
      Top             =   4080
      Value           =   1  'Checked
      Width           =   3255
   End
   Begin VB.CommandButton cmdClear 
      Caption         =   "clear picture box manually"
      Height          =   855
      Left            =   120
      TabIndex        =   11
      Top             =   7320
      Width           =   3255
   End
   Begin VB.CheckBox chkAutoPaint 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      Caption         =   "auto-paint to picture box"
      ForeColor       =   &H80000008&
      Height          =   375
      Left            =   240
      TabIndex        =   10
      Top             =   4440
      Value           =   1  'Checked
      Width           =   3135
   End
   Begin VB.HScrollBar scrOpacity 
      Height          =   300
      Left            =   240
      Max             =   100
      Min             =   1
      TabIndex        =   8
      Top             =   6840
      Value           =   100
      Width           =   2535
   End
   Begin VB.HScrollBar scrZoom 
      Height          =   300
      Left            =   240
      Max             =   300
      Min             =   1
      TabIndex        =   5
      Top             =   6120
      Value           =   100
      Width           =   2535
   End
   Begin VB.FileListBox lstImages 
      Appearance      =   0  'Flat
      Height          =   3090
      Left            =   120
      TabIndex        =   2
      Top             =   510
      Width           =   3255
   End
   Begin VB.PictureBox picDemo 
      Appearance      =   0  'Flat
      AutoRedraw      =   -1  'True
      BackColor       =   &H80000005&
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H80000008&
      Height          =   7215
      Left            =   3600
      ScaleHeight     =   479
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   727
      TabIndex        =   0
      Top             =   480
      Width           =   10935
   End
   Begin VB.Label lblSVGInfo 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "info on the active SVG will appear here"
      Height          =   375
      Left            =   3600
      TabIndex        =   16
      Top             =   7800
      Width           =   10935
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "when clicking on picture box:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   5
      Left            =   120
      TabIndex        =   15
      Top             =   5400
      Width           =   3255
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "when loading a new image:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   4
      Left            =   120
      TabIndex        =   14
      Top             =   3720
      Width           =   3255
   End
   Begin VB.Label lblOpacity 
      BackStyle       =   0  'Transparent
      Caption         =   "100%"
      Height          =   300
      Left            =   2880
      TabIndex        =   9
      Top             =   6840
      Width           =   615
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "custom opacity:"
      Height          =   375
      Index           =   3
      Left            =   240
      TabIndex        =   7
      Top             =   6480
      Width           =   3255
   End
   Begin VB.Label lblZoom 
      BackStyle       =   0  'Transparent
      Caption         =   "100%"
      Height          =   300
      Left            =   2880
      TabIndex        =   6
      Top             =   6120
      Width           =   615
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "custom zoom:"
      Height          =   375
      Index           =   2
      Left            =   240
      TabIndex        =   4
      Top             =   5760
      Width           =   3255
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "click the picture box to draw the selected SVG at that mouse position (using the zoom and opacity settings on the left):"
      Height          =   375
      Index           =   1
      Left            =   3600
      TabIndex        =   3
      Top             =   120
      Width           =   10935
   End
   Begin VB.Label lblTitle 
      BackStyle       =   0  'Transparent
      Caption         =   "sample SVG images:"
      BeginProperty Font 
         Name            =   "Segoe UI"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3255
   End
End
Attribute VB_Name = "frmDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***************************************************************************
'SVG Image Demo (via resvg) for VB6
'VB6 portion Copyright 2022 by Tanner Helland; see LICENSE.md and resvg-LICENSE.md for additional details
'Created: 28/February/22
'Last updated: 15/March/22
'Last update: wrap up initial build
'
'Please see the svgSupport module for full details on resvg, the third-party SVG library that makes
' this project possible.  The svgSupport module also details the license terms for resvg.  Please do not
' use this project until you familiarize yourself with the terms of that license.
'
'This small sample project demonstrates how to load and display SVG images in VB6.  Several sampel SVG
' images are included.
'
'SVG support via this project is extremely simple.  When your project starts, add the following line of code
' somewhere in program initialization:
'
' svgSupport.StartSVGSupport "C:\[path-to-resvg-folder]\resvg.dll"
'
'Obviously you will need to provide a valid path to resvg.dll.  This project ships with resvg in its
' App.Path folder, and you can look below to see how this works.
'
'Once SVG support is available, you can create as many individual SVG instances as you want using the
' svgSupport.LoadSVG_FromFile() function.  This function returns initialized instances of the svgImage class.
' Each class instance will manage associated SVG resources for you, meaning you can load an SVG once and
' then render it an infinite amount of times (at whatever sizes you want).
'
'Before your program exits, you need to stop the SVG engine using another call to the svgSupport module:
'
' svgSupport.StopSVGSupport
'
'This will release a number of shared SVG resources and unload the resvg library.  Importantly, you *MUST*
' ensure all svgImage instances are freed *BEFORE* calling this function.  (As you can imagine, unloading
' resvg before releasing classes that rely on it is a bad idea.)
'
'That's it!  Have fun!
'
'Please submit bug reports, feedback, etc. at GitHub:
' https://github.com/tannerhelland/vb6-svg
'
'Unless otherwise noted, all VB6 source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file in the root project directory.
'
'resvg is Copyright 2022 by Yevhenii Reizner.  It is used here under its original MPL-2 license.
' Full license details are available in the resvg-LICENSE.md file in the root project directory.
'
'***************************************************************************

Option Explicit

'The currently loaded SVG image.  We store it at module level, so that the user can "click" to draw it
' to arbitrary positions on the destination picture box.
Private m_activeSVG As svgImage

Private Sub cmdClear_Click()
    picDemo.Picture = LoadPicture(vbNullString)
    picDemo.Refresh
End Sub

Private Sub Form_Load()
    
    'BEFORE DOING ANYTHING ELSE, you MUST initialize the SVG library.  This involves
    ' telling the module where to find resvg.dll (for this demo project, that's App.Path).
    svgSupport.StartSVGSupport App.Path
    
    'Populate the list of sample images.  resvg supports plain svg and compressed svg (svgz).
    lstImages.Path = App.Path & "\images\"
    lstImages.Pattern = "*.svg;*.svgz"
    lstImages.ListIndex = 9
    
    'Sync scrollbar and label UI elements
    SynchronizeScrollLabels
    
End Sub

Private Sub Form_Unload(Cancel As Integer)

    'BEFORE YOUR APP EXITS, you MUST free the SVG library.
    ' BEFORE FREEING THE LIBRARY, you MUST ensure all SVG objects (svgImage instances) are freed.
    ' svgImage instances require SVG support to correct free their underlying SVG handles.
    Set m_activeSVG = Nothing
    svgSupport.StopSVGSupport

End Sub

'Click on the picture box to draw the active SVG with custom zoom and/or opacity values
Private Sub picDemo_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    
    'If our current SVG object contains a valid SVG, paint it at the specified position and/or zoom
    If (Not m_activeSVG Is Nothing) Then
        
        'Let's center the SVG at the requested draw position.  To do this, we offset the SVG by
        ' 1/2 of its width and height.
        Dim svgWidth As Long, svgHeight As Long
        svgWidth = m_activeSVG.GetDefaultWidth * (scrZoom.Value / 100!)
        svgHeight = m_activeSVG.GetDefaultHeight * (scrZoom.Value / 100!)
        
        m_activeSVG.DrawSVGtoDC picDemo.hDC, Int(x) - (svgWidth \ 2), Int(y) - (svgHeight \ 2), svgWidth, svgHeight, scrOpacity.Value
        
        'These lines are only required if the destination picture box .AutoRedraw = TRUE.  They will
        ' reflect any non-standard drawing changes (via GDI functions, for example) to the screen.
        picDemo.Picture = picDemo.Image
        picDemo.Refresh
        
    End If
    
End Sub

Private Sub scrOpacity_Change()
    SynchronizeScrollLabels
End Sub

Private Sub scrOpacity_Scroll()
    SynchronizeScrollLabels
End Sub

Private Sub scrZoom_Change()
    SynchronizeScrollLabels
End Sub

Private Sub scrZoom_Scroll()
    SynchronizeScrollLabels
End Sub

'Display the selected image.  Easy as can be!
Private Sub lstImages_Click()
    
    'Clear the picture box if the corresponding check box is set
    If CBool(chkAutoClear.Value) Then picDemo.Picture = LoadPicture(vbNullString)
    
    'Attempt to load the selected SVG
    If svgSupport.LoadSVG_FromFile(lstImages.Path & "\" & lstImages.FileName, m_activeSVG) Then
        
        'If the user wants us to auto-paint newly loaded images, do so now.  Note that we do not
        ' apply the user's selected zoom or opacity values - those are applied when clicking the
        ' picture box, only.
        If (m_activeSVG.HasSVG()) And CBool(chkAutoPaint.Value) Then
            
            'Retrieve the default dimensions from the SVG
            Dim fitWidth As Long, fitHeight As Long
            fitWidth = m_activeSVG.GetDefaultWidth
            fitHeight = m_activeSVG.GetDefaultHeight
            
            'The user can choose to have us "auto-fit" the image to the display area.  (This makes
            ' for a great demonstration of the "infinite resizing" capabilities of vector graphics.)
            If CBool(chkFit.Value) Then
                
                'Fit the vertical dimension first
                fitWidth = fitWidth * (picDemo.ScaleHeight / fitHeight)
                fitHeight = picDemo.ScaleHeight
                
                'If our scaling made the image too wide to fit in the picture box,
                ' fit the width as well.
                If (fitWidth > picDemo.ScaleWidth) Then
                    fitHeight = fitHeight * (picDemo.ScaleWidth / fitWidth)
                    fitWidth = picDemo.ScaleWidth
                End If
                
            End If
            
            'Draw the SVG to the picture box.  Note that you can paint to any (x, y) position and
            ' any arbitrary size.  (You can also supply custom opacity, but we don't do that here -
            ' check out the picture box's _MouseDown event to see that feature in action.)
            m_activeSVG.DrawSVGtoDC picDemo.hDC, _
                                    (picDemo.ScaleWidth - fitWidth) \ 2, _
                                    (picDemo.ScaleHeight - fitHeight) \ 2, _
                                    fitWidth, _
                                    fitHeight
            
            'These lines are only required if the destination picture box .AutoRedraw = TRUE.  They will
            ' reflect any non-standard drawing changes (via SVG libraries, for example) to the screen.
            picDemo.Picture = picDemo.Image
            picDemo.Refresh
            
        End If
        
        'Display the SVG's default size in the bottom label
        lblSVGInfo.Caption = "this SVG has a default size of " & m_activeSVG.GetDefaultWidth & " x " & m_activeSVG.GetDefaultHeight
        
    Else
        Set m_activeSVG = Nothing
        Debug.Print "Failed to load " & lstImages.Path & "\" & lstImages.FileName
    End If
    
End Sub

Private Sub SynchronizeScrollLabels()
    lblZoom.Caption = Format$(scrZoom.Value / 100, "0%")
    lblOpacity.Caption = Format$(scrOpacity.Value / 100, "0%")
End Sub
