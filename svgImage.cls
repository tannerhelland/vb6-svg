VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "svgImage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'***************************************************************************
'SVG Image Container (via resvg) for VB6
'VB6 portion Copyright 2022 by Tanner Helland; see LICENSE.md and resvg-LICENSE.md for additional details
'Created: 28/February/22
'Last updated: 15/March/22
'Last update: wrap up initial build
'
'Please see the svgSupport module for full details on resvg, the third-party SVG library that makes
' this project possible.
'
'This class holds a single SVG instance.  It's designed to let you load an arbitrary number of SVGs
' without worrying about resource management.
'
'Instances of this class should be auto-initialized using the svgSupport.LoadSVG_FromFile() function.
'
'You must observe two critical rules while using this class:
' 1) Instances of this class must NOT be instantiated until AFTER you call svgSupport.StartSVGSupport().
' 2) Instances of this class must be freed BEFORE you call svgSupport.StopSVGSupport().
'
'AS LONG AS YOU ADHERE TO THE TWO RULES ABOVE, you don't need to worry about freeing any SVG-related
' resources.  This class handles all of that for you.
'
'After successfully loading an SVG image, you can draw it to any arbitrary hDC (e.g. PictureBox.hDC
' or Form.hDC) using the DrawSVGToDC function.  This function also allows you specify custom (x, y)
' offsets, custom width/height (with "lossless" vector scaling), and custom opacity.
'
'See the attached frmDemo UI for a full demonstration.
'
'Please submit bug reports, feedback, etc. at GitHub:
' https://github.com/tannerhelland/vb6-svg
'
'Unless otherwise noted, all VB6 source code in this file is shared under a simplified BSD license.
' Full license details are available in the LICENSE.md file in the root project directory.
'
'resvg is Copyright 2024 by the resvg authors.  It is used here under its original MPL-2 license.
' Full license details are available in the resvg-LICENSE.md file in the root project directory.
'***************************************************************************

Option Explicit

'Active SVG DOM handle, if any.  *MUST* be freed before exiting (or replacing with a new handle)
' or it will leak.
Private m_hTree As Long

'Once we've been assigned a handle, we also track the "default" width/height supplied by the SVG (viewbox)
Private m_defaultWidth As Long, m_defaultHeight As Long

'Draw the SVG to an arbitrary DC.  Rules:
' 1) All measurements must be in pixels.
' 2) If you are drawing to a VB object with AutoRedraw set to TRUE, you must manually refresh the image
'    after rendering (see frmDemo for details on this).
' 3) Pass zero for the optional width or height value to use the default width/height value specified by
'    the SVG viewbox.
' 4) Opacity is specified as a floating-point value on the scale [0, 100], where 100 = "full opacity".
Friend Function DrawSVGtoDC(ByVal dstDC As Long, Optional ByVal dstX As Long = 0, Optional ByVal dstY As Long = 0, Optional ByVal dstWidth As Long = 0, Optional ByVal dstHeight As Long = 0, Optional ByVal dstOpacity As Single = 100!) As Boolean
    
    Const FUNC_NAME As String = "DrawSVGtoDC"
    
    'Validate inputs
    If (dstDC = 0) Then
        InternalError FUNC_NAME, "null DC"
        DrawSVGtoDC = False
        Exit Function
    End If
    
    If (dstOpacity <= 0!) Or (dstOpacity > 100!) Then
        InternalError FUNC_NAME, "bad opacity value"
        DrawSVGtoDC = False
        Exit Function
    End If
    
    If (Not Me.HasSVG) Then
        InternalError FUNC_NAME, "you have to load an SVG before drawing it!"
        DrawSVGtoDC = False
        Exit Function
    End If
    
    If (m_defaultWidth <= 0) Or (m_defaultHeight <= 0) Then
        InternalError FUNC_NAME, "SVG has no width/height"
        DrawSVGtoDC = 0
        Exit Function
    End If
    
    'svgSupport handles the actual rendering
    DrawSVGtoDC = svgSupport.INTERNAL_DrawTreeToDC(dstDC, Me, dstX, dstY, dstWidth, dstHeight, dstOpacity)
    
End Function

'Query embedded SVG width/height
Friend Function GetDefaultWidth() As Long
    GetDefaultWidth = m_defaultWidth
End Function

Friend Function GetDefaultHeight() As Long
    GetDefaultHeight = m_defaultHeight
End Function

'Returns TRUE if this class has been assigned a valid SVG DOM handle
Friend Function HasSVG() As Boolean
    HasSVG = (m_hTree <> 0)
End Function

'Free our SVG handle (if any) and reset all class behavior to its default state
Friend Sub Reset()
    Free_hSVG
End Sub

'All INTERNAL-prefixed functions must only interact with the svgSupport module.
' Do not call them directly or you will break things.
Friend Function INTERNAL_GetSVGHandle() As Long
    INTERNAL_GetSVGHandle = m_hTree
End Function

Friend Sub INTERNAL_SetSVGHandle(ByVal srcHandle As Long, ByVal defaultWidth As Long, ByVal defaultHeight As Long)
    
    'If we already have an SVG handle, kill it
    If (m_hTree <> 0) Then Free_hSVG
    m_hTree = srcHandle
    
    If (m_hTree <> 0) Then
        m_defaultWidth = defaultWidth
        m_defaultHeight = defaultHeight
    Else
        Free_hSVG
    End If
    
End Sub

Private Sub Class_Initialize()
    'Nothing is allocated on initialization
End Sub

Private Sub Class_Terminate()
    
    'ALWAYS free our DOM handle before exiting
    Free_hSVG
    
End Sub

Private Sub Free_hSVG()
    If (m_hTree <> 0) Then svgSupport.INTERNAL_FreeTree m_hTree
    m_hTree = 0
    m_defaultWidth = 0
    m_defaultHeight = 0
End Sub

Private Sub InternalError(ByRef funcName As String, ByRef errMsg As String)
    Debug.Print "WARNING! svgImage." & funcName & "() error: " & errMsg
End Sub

