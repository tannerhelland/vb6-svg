Attribute VB_Name = "svgSupport"
'***************************************************************************
'resvg Library Interface (SVG import) for VB6
'VB6 portion Copyright 2022 by Tanner Helland; see LICENSE.md and resvg-LICENSE.md for additional details
'Created: 28/February/22
'Last updated: 15/March/22
'Last update: wrap up initial build
'
'This project demonstrates full SVG support in VB6 c/o the third-party resvg library.  It is designed
' to be leak-free and very simple to use.
'
'That said, you must observe TWO RULES for this project to remain leak-free:
'
'1) You MUST call svgSupport.StartSVGSupport() before using any SVG features.  This function asks
'   you to pass the folder where your copy of resvg.dll resides (perhaps App.Path or similar).
'   Until you call this function, no SVG features will work.
'2) You MUST call svgSupport.StopSVGSupport() before your app ends.  Before calling this function,
'   you MUST ensure that all svgImage instances have fallen out of scope (or manually set to Nothing).
'   svgImage instances require resvg to be active, since resvg handles actual resource allocation/disposal.
'   As long as all svgImage instances are released BEFORE calling the StopSVGSupport() function,
'   this project will not leak memory or resources.
'
'To get started, read the comments on svgSupport.LoadSVG_FromFile(), below, or simply hit F5 and play
' around with the attached demo UI (frmDemo).
'
'ABOUT resvg:
'
'Per its documentation (available at https://github.com/RazrFalcon/resvg), resvg is...
'
'"...an SVG rendering library. It can be used as a Rust library, as a C library and as a CLI application
' to render static SVG files. The core idea is to make a fast, small, portable SVG library with an aim to
' support the whole SVG spec."
'
'Yevhenii Reizner is the author of resvg.  resvg is MPL-licensed and actively maintained.  If you don't
' know what the terms of the MPL license include, please study the provided resvg-LICENSE.md file to
' ensure the license is compatible with your intended use-case.  A full copy of the resvg license is
' also available online: https://github.com/RazrFalcon/resvg/blob/master/LICENSE.txt
'
'BUILDING resvg:
'
'The copy of resvg.dll that ships with this project is based on the 0.22.0 release and built against
' the i686-pc-windows-msvc rust target (for Win Vista+ support).  It *must* be hand-edited to export
' stdcall funcs. (You might be tempted to just use cdecl via DispCallFunc, but some resvg functions
' return custom types that won't work with DispCallFunc - so manually building against stdcall is
' mandatory.)  Note that some function decs must also be rewritten to pass UDTs as references instead
' of values, as required by VB6.  This makes the build somewhat non-standard, so I do not recommend
' building the library yourself without carefully reviewing the matching VB6 declarations provided below.
'
'Finally, you can read all about resvg default settings and how to customize them from the C-api header
' available at GitHub: https://github.com/RazrFalcon/resvg/blob/master/c-api/resvg.h
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

'Information on individual resvg calls can be saved to the debug log via this constant;
' consider DISABLING in production builds as it will display (potentially) a LOT of info.
Private Const SVG_DEBUG_VERBOSE As Boolean = False

'Enable to support embedded SVG text (inc. fonts).  Disable for better performance, but note
' that SVG text objects will only render using the default font (likely Times New Roman,
' although you can change this via the resvg C API).
Private Const SVG_ENABLE_TEXT As Boolean = True

Private Enum resvg_result
    'Everything is ok.
    RESVG_OK = 0
    'Only UTF-8 content are supported.
    RESVG_ERROR_NOT_AN_UTF8_STR = 1
    'Failed to open the provided file.
    RESVG_ERROR_FILE_OPEN_FAILED = 2
    'Compressed SVG must use the GZip algorithm.
    RESVG_ERROR_MALFORMED_GZIP = 3
    'We do not allow SVG with more than 1_000_000 elements for security reasons.
    RESVG_ERROR_ELEMENTS_LIMIT_REACHED = 4
    'SVG doesn't have a valid size.
    '   (Occurs when width and/or height are <= 0.)
    '   (Also occurs if width, height and viewBox are not set.)
    RESVG_ERROR_INVALID_SIZE = 5
    'Failed to parse an SVG data.
    RESVG_ERROR_PARSING_FAILED = 6
End Enum

#If False Then
    Private Const RESVG_OK = 0, RESVG_ERROR_NOT_AN_UTF8_STR = 1, RESVG_ERROR_FILE_OPEN_FAILED = 2, RESVG_ERROR_MALFORMED_GZIP = 3, RESVG_ERROR_ELEMENTS_LIMIT_REACHED = 4, RESVG_ERROR_INVALID_SIZE = 5, RESVG_ERROR_PARSING_FAILED = 6
#End If

'A "fit to" type.
' (All types produce proportional scaling.)
Private Enum resvg_fit_to_type
    'Use an original image size.
    RESVG_FIT_TO_TYPE_ORIGINAL
    'Fit an image to a specified width.
    RESVG_FIT_TO_TYPE_WIDTH
    'Fit an image to a specified height.
    RESVG_FIT_TO_TYPE_HEIGHT
    'Zoom an image using scaling factor.
    RESVG_FIT_TO_TYPE_ZOOM
End Enum

#If False Then
    Private Const RESVG_FIT_TO_TYPE_ORIGINAL = 0, RESVG_FIT_TO_TYPE_WIDTH = 0, RESVG_FIT_TO_TYPE_HEIGHT = 0, RESVG_FIT_TO_TYPE_ZOOM = 0
#End If

'An image rendering method.
Private Enum resvg_image_rendering
    RESVG_IMAGE_RENDERING_OPTIMIZE_QUALITY
    RESVG_IMAGE_RENDERING_OPTIMIZE_SPEED
End Enum

#If False Then
    Private Const RESVG_IMAGE_RENDERING_OPTIMIZE_QUALITY = 0, RESVG_IMAGE_RENDERING_OPTIMIZE_SPEED = 0
#End If

'A shape rendering method.
Private Enum resvg_shape_rendering
    RESVG_SHAPE_RENDERING_OPTIMIZE_SPEED
    RESVG_SHAPE_RENDERING_CRISP_EDGES
    RESVG_SHAPE_RENDERING_GEOMETRIC_PRECISION
End Enum

#If False Then
    Private Const RESVG_SHAPE_RENDERING_OPTIMIZE_SPEED = 0, RESVG_SHAPE_RENDERING_CRISP_EDGES = 0, RESVG_SHAPE_RENDERING_GEOMETRIC_PRECISION = 0
#End If

'A text rendering method.
Private Enum resvg_text_rendering
    RESVG_TEXT_RENDERING_OPTIMIZE_SPEED
    RESVG_TEXT_RENDERING_OPTIMIZE_LEGIBILITY
    RESVG_TEXT_RENDERING_GEOMETRIC_PRECISION
End Enum

#If False Then
    Private Const RESVG_TEXT_RENDERING_OPTIMIZE_SPEED = 0, RESVG_TEXT_RENDERING_OPTIMIZE_LEGIBILITY = 0, RESVG_TEXT_RENDERING_GEOMETRIC_PRECISION = 0
#End If

'A 2D transform representation.
Private Type resvg_transform
    a As Double
    b As Double
    c As Double
    d As Double
    e As Double
    f As Double
End Type

'A size representation.
' (Width and height are guaranteed to be > 0.)
Private Type resvg_size
    svg_width As Double
    svg_height As Double
End Type

'A rectangle representation.
' (Width *and* height are guarantee to be > 0.)
Private Type resvg_rect
    x As Double
    y As Double
    Width As Double
    Height As Double
End Type

'A path bbox representation.
' (Width *or* height are guarantee to be > 0.)
Private Type resvg_path_bbox
    x As Double
    y As Double
    Width As Double
    Height As Double
End Type

'A "fit to" property.
Private Type resvg_fit_to
    'A fit type.
    fit_type As resvg_fit_to_type
    'Fit to value
    '* Not used by RESVG_FIT_TO_ORIGINAL.
    '* Must be >= 1 for RESVG_FIT_TO_WIDTH and RESVG_FIT_TO_HEIGHT.
    '* Must be > 0 for RESVG_FIT_TO_ZOOM.
    fit_value As Single
End Type

Private Declare Function resvg_transform_identity Lib "resvg" () As resvg_transform
Private Declare Sub resvg_init_log Lib "resvg" ()
Private Declare Function resvg_options_create Lib "resvg" () As Long
Private Declare Sub resvg_options_set_resources_dir Lib "resvg" (ByVal resvg_options As Long, ByVal ptrToConstUtf8Family As Long)
Private Declare Sub resvg_options_set_dpi Lib "resvg" (ByVal resvg_options As Long, ByVal newDPI As Double)
Private Declare Sub resvg_options_set_font_family Lib "resvg" (ByVal resvg_options As Long, ByVal ptrToConstUtf8Family As Long)
Private Declare Sub resvg_options_set_font_size Lib "resvg" (ByVal resvg_options As Long, ByVal newSize As Double)
Private Declare Sub resvg_options_set_serif_family Lib "resvg" (ByVal resvg_options As Long, ByVal ptrToConstUtf8Family As Long)
Private Declare Sub resvg_options_set_sans_serif_family Lib "resvg" (ByVal resvg_options As Long, ByVal ptrToConstUtf8Family As Long)
Private Declare Sub resvg_options_set_cursive_family Lib "resvg" (ByVal resvg_options As Long, ByVal ptrToConstUtf8Family As Long)
Private Declare Sub resvg_options_set_fantasy_family Lib "resvg" (ByVal resvg_options As Long, ByVal ptrToConstUtf8Family As Long)
Private Declare Sub resvg_options_set_monospace_family Lib "resvg" (ByVal resvg_options As Long, ByVal ptrToConstUtf8Family As Long)
Private Declare Sub resvg_options_set_languages Lib "resvg" (ByVal resvg_options As Long, ByVal ptrToConstUtf8Languages As Long)
Private Declare Sub resvg_options_set_shape_rendering_mode Lib "resvg" (ByVal resvg_options As Long, ByVal newMode As resvg_shape_rendering)
Private Declare Sub resvg_options_set_text_rendering_mode Lib "resvg" (ByVal resvg_options As Long, ByVal newMode As resvg_text_rendering)
Private Declare Sub resvg_options_set_image_rendering_mode Lib "resvg" (ByVal resvg_options As Long, ByVal newMode As resvg_image_rendering)
Private Declare Sub resvg_options_set_keep_named_groups Lib "resvg" (ByVal resvg_options As Long, ByVal keepBool As Long)
Private Declare Sub resvg_options_load_font_data Lib "resvg" (ByVal resvg_options As Long, ByVal ptrToData As Long, ByVal sizeOfData As Long)
Private Declare Function resvg_options_load_font_file Lib "resvg" (ByVal resvg_options As Long, ByVal ptrToConstUtf8FilePath As Long) As resvg_result
Private Declare Sub resvg_options_load_system_fonts Lib "resvg" (ByVal resvg_options As Long)
Private Declare Sub resvg_options_destroy Lib "resvg" (ByVal resvg_options As Long)
Private Declare Function resvg_parse_tree_from_file Lib "resvg" (ByVal ptrToConstUtf8FilePath As Long, ByVal resvg_options As Long, ByRef resvg_render_tree As Long) As Long
Private Declare Function resvg_parse_tree_from_data Lib "resvg" (ByVal ptrToData As Long, ByVal sizeOfData As Long, ByVal resvg_options As Long, ByRef resvg_render_tree As Long) As Long
Private Declare Function resvg_is_image_empty Lib "resvg" (ByVal resvg_render_tree As Long) As Long
Private Declare Function resvg_get_image_size Lib "resvg" (ByVal resvg_render_tree As Long) As resvg_size
Private Declare Function resvg_get_image_viewbox Lib "resvg" (ByVal resvg_render_tree As Long) As resvg_rect
Private Declare Function resvg_get_imgae_bbox Lib "resvg" (ByVal resvg_render_tree As Long, ByRef dst_resvg_rect As resvg_rect) As Long
Private Declare Function resvg_node_exists Lib "resvg" (ByVal resvg_render_tree As Long, ByVal ptrToConstUtf8ID As Long) As Long
Private Declare Function resvg_get_node_transform Lib "resvg" (ByVal resvg_render_tree As Long, ByVal ptrToConstUtf8ID As Long, ByRef dst_resvg_transform As resvg_transform) As Long
Private Declare Function resvg_get_node_bbox Lib "resvg" (ByVal resvg_render_tree As Long, ByVal ptrToConstUtf8ID As Long, ByRef dst_resvg_path_bbox As resvg_path_bbox) As Long
Private Declare Sub resvg_tree_destroy Lib "resvg" (ByVal resvg_render_tree As Long)
Private Declare Sub resvg_render Lib "resvg" (ByVal resvg_render_tree As Long, ByRef fit_to As resvg_fit_to, ByRef srcTransform As resvg_transform, ByVal surfaceWidth As Long, ByVal surfaceHeight As Long, ByVal ptrToSurface As Long)
Private Declare Sub resvg_render_node Lib "resvg" (ByVal resvg_render_tree As Long, ByVal ptrToConstUtf8ID As Long, ByVal fit_to As resvg_fit_to, ByVal srcTransform As resvg_transform, ByVal surfaceWidth As Long, ByVal surfaceHeight As Long, ByVal ptrToSurface As Long)

'Generic WAPI support functions
Private Declare Sub CopyMemoryStrict Lib "kernel32" Alias "RtlMoveMemory" (ByVal lpDst As Long, ByVal lpSrc As Long, ByVal byteLength As Long)
Private Declare Function FreeLibrary Lib "kernel32" (ByVal hLibModule As Long) As Long
Private Declare Function LoadLibraryW Lib "kernel32" (ByVal lpLibFileName As Long) As Long
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal dstCodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpDefaultChar As Long, ByVal lpUsedDefaultChar As Long) As Long
Private Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (ByVal dstPointer As Long, ByVal numOfBytes As Long)

'GDI interop
Private Type BITMAPINFOHEADER
    Size As Long
    Width As Long
    Height As Long
    Planes As Integer
    BitCount As Integer
    Compression As Long
    ImageSize As Long
    xPelsPerMeter As Long
    yPelsPerMeter As Long
    ColorUsed As Long
    ColorImportant As Long
End Type

Private Declare Function AlphaBlend Lib "gdi32" Alias "GdiAlphaBlend" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal WidthSrc As Long, ByVal HeightSrc As Long, ByVal blendFunct As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32" (ByVal hDC As Long, ByRef lpBitsInfo As BITMAPINFOHEADER, ByVal wUsage As Long, ByRef lpBits As Long, ByVal hSection As Long, ByVal dwOffset As Long) As Long
Private Declare Function DeleteDC Lib "gdi32" (ByVal hDC As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hDC As Long, ByVal hObject As Long) As Long

'Used for creating arrays that point at arbitrary data
Private Type SafeArray1D
    cDims      As Integer
    fFeatures  As Integer
    cbElements As Long
    cLocks     As Long
    pvData     As Long
    cElements As Long
    lBound   As Long
End Type

Private Declare Sub PutMem4 Lib "msvbvm60" (ByVal ptrDst As Long, ByVal newValue As Long)
Private Declare Function VarPtrArray Lib "msvbvm60" Alias "VarPtr" (ByRef srcArray() As Any) As Long

'Debug tracker for active SVG DOM instances.  If this is not zero when you stop the SVG engine,
' you have leaked handles.  Shame on you.  (Or shame on me, if it was my fault - either way,
' this will be reported in the debug window and you should figure out what's wrong.)
Private m_ActiveTree As Long

'A single persistent SVG options handle is maintained for the life of a session.
' (Initializing this object is expensive because it needs to scan system fonts.)
Private m_Options As Long

'Library handle will be non-zero if all required dll(s) are available;
' you can also forcibly override the "availability" state by setting m_LibAvailable to FALSE.
' (This effectively disables run-time support in the UI.)
Private m_libHandle As Long, m_libAvailable As Boolean

'Because resvg only supports 32-bit RGBA surfaces, and VB6 does not support native 32-bit RGBA surfaces,
' we must create our own intermediary surface.  GDI works fine for this, because we can simply AlphaBlend
' the rendered result to any arbitrary DC with minimal fuss.
Private m_tmpDIB As Long, m_tmpDIBHeader As BITMAPINFOHEADER, m_tmpDIBBits As Long
Private m_tmpDC As Long, m_tmpOldObject As Long

'To improve performance, we only re-render the full SVG when necessary.
' (Basically, if a draw request matches the last draw request, we can reuse our existing GDI buffer.)
Private m_lastSVG As Long, m_lastWidth As Long, m_lastHeight As Long

'Helper function to report if SVG support has been properly initialized
Public Function IsSVGEnabled() As Boolean
    IsSVGEnabled = m_libAvailable
End Function

'You MUST call this function before attempting to load any SVGs.
'
'Note also that this function will fail on Windows XP.  You must use Vista or later.
Public Function StartSVGSupport(ByRef pathToDLLFolder As String) As Boolean
    
    Const FUNC_NAME As String = "StartSVGSupport"
    
    'Feel free to hard-code this path as relevant to your use-case.  The idea with passing
    ' the folder that contains this library is that you can ship the library wherever you want
    ' (a subfolder within your app folder, perhaps).
    Dim strLibPath As String
    strLibPath = pathToDLLFolder
    If (Right$(strLibPath, 1) <> "\") And (Right$(strLibPath, 1) <> "/") Then strLibPath = strLibPath & "\"
    strLibPath = strLibPath & "resvg.dll"
    
    m_libHandle = LoadLibraryW(StrPtr(strLibPath))
    m_libAvailable = (m_libHandle <> 0)
    StartSVGSupport = m_libAvailable
        
    If SVG_DEBUG_VERBOSE Then
        If (Not StartSVGSupport) Then InternalError FUNC_NAME, "LoadLibraryW failed to load resvg.  Last DLL error: " & Err.LastDllError
    End If
    
End Function

'You MUST call this function...
' 1) BEFORE your program ends, and...
' 2) AFTER all imgSVG class instances are freed.  (Those classes cannot be safely freed without
'    active SVG support, because we need to free opaque SVG handles using resvg.)
Public Sub StopSVGSupport()
    
    Const FUNC_NAME As String = "StopSVGSupport"
    
    'Free the persistent options handle, if it exists
    If (m_Options <> 0) Then
        resvg_options_destroy m_Options
        m_Options = 0
    End If
    
    'Free the library itself
    If (m_libHandle <> 0) Then
        FreeLibrary m_libHandle
        m_libHandle = 0
    End If
    
    'Free any GDI objects we created
    If (m_tmpDC <> 0) Then
        If (m_tmpDIB <> 0) Then SelectObject m_tmpDC, m_tmpOldObject
        DeleteDC m_tmpDC
        m_tmpDC = 0
    End If
    
    If (m_tmpDIB <> 0) Then
        DeleteObject m_tmpDIB
        m_tmpDIB = 0
    End If
    
    With m_tmpDIBHeader
        .Width = 0
        .Height = 0
        .ImageSize = 0
    End With
    
    m_libAvailable = False
    
End Sub

'Load (import) functions follow

'Load an SVG from file.  You *must* supply a valid file path and you *must* supply a destination svgImage object.
' Returns TRUE if successful; FALSE otherwise.  FALSE returns will print a failure reason to the debug window.
Public Function LoadSVG_FromFile(ByRef srcFile As String, ByRef dstImage As svgImage) As Boolean
    
    Const FUNC_NAME As String = "LoadSVG_FromFile"
    
    LoadSVG_FromFile = False
    
    If SVG_DEBUG_VERBOSE Then Debug.Print "Attempting to load SVG: " & srcFile
    
    'Ensure a blank resvg options object exists.  This "options" handle specifies all load-time behavior.
    ' Creating it is potentially expensive (especially if you want full font support, because the system
    ' font collection must be scanned), but we only have to do it *once* - then we can reuse it for all
    ' SVG requests that follow.
    If (m_Options = 0) Then
        
        'Retrieve a "default settings" handle
        m_Options = resvg_options_create()
        
        'If you *don't* care about text support, you can comment out this line of code for
        ' a performance boost.
        If SVG_ENABLE_TEXT Then
            resvg_options_load_system_fonts m_Options
            If SVG_DEBUG_VERBOSE Then Debug.Print "(Note: this session supports embedded SVG text)"
        End If
        
    End If
        
    'Pre-set any other options here, as desired.
    '
    'For a full list of default settings (and how to modify them), consult resvg.h:
    ' https://github.com/RazrFalcon/resvg/blob/master/c-api/resvg.h
    
    'Ensure a valid handle object exists before attempting further loading.
    If (m_Options = 0) Then
        If SVG_DEBUG_VERBOSE Then Debug.Print "WARNING: couldn't create master SVG options handle."
        LoadSVG_FromFile = False
        Exit Function
    End If
    
    'Armed with our new options handle, we can now attempt to build a DOM for the target SVG.
    
    '(All resvg calls operate on UTF-8 strings, so you will see many conversions from BSTRs
    ' to UTF8 bytestreams throughout this module.)
    'Create a blank resvg render tree pointer, and note that this is the first call where we get
    ' an actual success/fail return
    Dim svgResult As resvg_result, svgTree As Long
    
    Dim utf8path() As Byte, utf8Len As Long
    If UTF8FromString(srcFile, utf8path, utf8Len) Then
        svgResult = resvg_parse_tree_from_file(VarPtr(utf8path(0)), m_Options, svgTree)
    Else
        InternalError FUNC_NAME, "bad input file; couldn't generate UTF-8 equivalent"
    End If
    
    If (svgResult = RESVG_OK) Then
        If SVG_DEBUG_VERBOSE Then Debug.Print "Successfully retrieved SVG tree: " & svgTree
    Else
        
        'Check error state.  Some errors may be recoverable.
        If (svgResult = RESVG_ERROR_PARSING_FAILED) Then
            
            'This error almost always means a malformed SVG.  You may be able to overcome the
            ' error by manually modifying the SVG data, then reloading the file.  This kind of
            ' detailed handling is left as an exercise for the user.
            
        '/failed for some other reason than a bad parse
        Else
            Debug.Print srcFile
            InternalError FUNC_NAME, "resvg error code (see top of svgSupport module for enum): " & svgResult
            LoadSVG_FromFile = False
            Exit Function
        End If
        
    End If
    
    'Failsafe check for non-zero handle
    If (svgTree <> 0) Then
        
        'We are now going to hand this persistent SVG DOM handle off to a child class.  That class
        ' will auto-free the handle when the class is freed.
        Set dstImage = New svgImage
        
        'Retrieve image size and convert to integers (to be used as surface dimensions)
        Dim imgSize As resvg_size
        imgSize = resvg_get_image_size(svgTree)
        
        Dim intWidth As Long, intHeight As Long
        intWidth = Int(imgSize.svg_width)
        intHeight = Int(imgSize.svg_height)
        If (intWidth < 1) Then intWidth = 1
        If (intHeight < 1) Then intHeight = 1
        
        'Assign these values to the target image instance
        dstImage.INTERNAL_SetSVGHandle svgTree, intWidth, intHeight
        LoadSVG_FromFile = True
        
    Else
        InternalError FUNC_NAME, "SVG load supposedly worked, but DOM handle is 0?  Please investigate."
        LoadSVG_FromFile = False
        Exit Function
    End If
    
End Function

'Public support functions follow.  Do *NOT* engage with these functions directly;
' only svgImage class instances should touch functions prefixed with INTERNAL.

'Draw an SVG object to an arbtrary hDC
Public Function INTERNAL_DrawTreeToDC(ByVal dstDC As Long, ByRef srcImage As svgImage, Optional ByVal dstX As Long = 0, Optional ByVal dstY As Long = 0, Optional ByVal dstWidth As Long = 0, Optional ByVal dstHeight As Long = 0, Optional ByVal svgOpacity As Single = 100!) As Boolean

    Const FUNC_NAME As String = "INTERNAL_DrawTreeToDC"
    
    'Because this function is explicitly marked as INTERNAL, we won't waste time validating inputs.
    ' (It's assumed the caller has already done this.)
    
    'resvg is an incredibly small but full-featured SVG library.  It achieves its tiny size by
    ' using an incredibly stripped-down copy of skia (called "tiny-skia": https://github.com/RazrFalcon/tiny-skia)
    ' as its canvas renderer.  Its minimalist rendering approach creates two problems for VB6 usage:
    ' 1) tiny-skia only renders to 32-bit surfaces.  Windows DCs make no guarantees about color-depth
    '    (and even if they did, VB6 doesn't play well with 32-bit surfaces).
    ' 2) tiny-skia only renders in RGBA order.  Windows surfaces are always BGRA.
    
    'To work around this, we manually create a 32-bit GDI surface (DIB) to receive the initial
    ' SVG render.  We then swizzle the red and blue channels, then apply a final GDI AlphaBlend
    ' at the requested opacity level.
    
    'This approach works well and has minimal overhead, but we must make some hard choices on
    '  resource usage vs performance.  There are basically two approaches we could take:
    ' 1) Let each svgImage class manage its own surface.  This would mean 1x DIB and DC per svgImage,
    '    but if SVG properties (like size) don't change, we can simply paint from the cached DIB and
    '    get incredibly good performance.
    ' 2) Share one surface across all svgImage instances.  This means just 1x DIB and DC for the entire
    '    SVG engine, but if different-size SVGs need to be painted, we have no choice but to re-create
    '    the shared DIB on every call.  This reduces performance but ensures minimal memory consumption.
    
    'Given VB6's typical usage scenarios, I have opted for option (2) in this module.  A single GDI
    ' surface (and associated DC) are shared across all SVG instances, which keeps this SVG engine
    ' incredibly lightweight.  For small SVG images (UI icons, etc) there is no meaningful performance
    ' hit to this approach.  If you do the correct thing and compile to native code with the
    ' Advanced Optimization "Remove Array Bounds Check" enabled, the GDI interop portion of SVG rendering
    ' is basically instantaneous, even on ancient hardware.
    
    'That said, if you need to render extremely large SVGs with wildly different sizes, you will benefit
    ' from switching to strategy (1) as described above.  To do so, simply move the GDI code from this
    ' module into the svgImage class, perform all size synchronization there, and only rebuild your GDI
    ' surface if the desired SVG size changes.  (Advanced users should just manage their own memory-mapped
    ' backing surface for the DIB - then you could reuse memory and have measurable perf impact for
    ' surface initialization.  Swizzling still requires manual intervention, however.)
    
    'Anyway, let's get on with GDI surface creation.
        
    'If this request is the same as the last render request we received (same SVG, same width, same height),
    ' we can reuse our existing GDI buffer as-is.  This improves performance.
    If (m_lastSVG <> srcImage.INTERNAL_GetSVGHandle()) Or (m_lastWidth <> dstWidth) Or (m_lastHeight <> dstHeight) Then
        
        'Unfortunately, this request is for a new SVG and/or dimension, which means we need a new GDI buffer
        ' to hold the intermediary result.
        
        'Ensure we have a valid DIB and DC ready to receive the SVG.
        EnsureGDISurfaceReady dstWidth, dstHeight
        
        'Make sure the GDI surface exists
        If (m_tmpDIB = 0) Or (m_tmpDIBBits = 0) Or (m_tmpDC = 0) Then
            InternalError FUNC_NAME, "one or more bad GDI objects"
            INTERNAL_DrawTreeToDC = False
            Exit Function
        End If
        
        'We can now perform the SVG render.  Note that we are obviously rendering to the temporary GDI object,
        ' *NOT* the destination DC.  (That will happen later.)
        
        'Specify fitting behavior (we always use original fit - you'll see why in a moment)
        Dim fitBehavior As resvg_fit_to
        fitBehavior.fit_type = RESVG_FIT_TO_TYPE_ORIGINAL
        fitBehavior.fit_value = 1!
        
        'If custom destination width/height is specified, we want to use the final transform matrix
        ' to apply the resize.
        Dim idMatrix As resvg_transform
        idMatrix = resvg_transform_identity()
        
        If (dstWidth <> srcImage.GetDefaultWidth()) Or (dstHeight <> srcImage.GetDefaultHeight()) Then
            
            'Scaling is required.  Populate the transform matrix accordingly.
            ' (This is just an affine transform matrix, but the order is weirdly written as:
            ' [a c e]
            ' [b d f])
            idMatrix.a = dstWidth / srcImage.GetDefaultWidth
            idMatrix.d = dstHeight / srcImage.GetDefaultHeight
            
            'Note that you could also apply rotation or skew here, if desired.
            
        End If
        
        'Render!
        resvg_render srcImage.INTERNAL_GetSVGHandle(), fitBehavior, idMatrix, dstWidth, dstHeight, m_tmpDIBBits
            
        'Before exiting, we need to swizzle red and blue channels.  The fastest way to do this
        ' (in VB6) is to "wrap" an array around the bits we just painted.  We can do this by
        ' building a SAFEARRAY descriptor (what VB uses to define an array), plugging in our
        ' GDI DIB pointer, then using WAPI to associate our custom SAFEARRAY descriptor with an
        ' actual array reference.  VB won't know the difference between this hand-crafted array
        ' and a normal VB6 array, but note that we *must* free our hand-crafted array before it
        ' exits scope or VB will try to free memory that it hasn't allocated.
        Dim imgPixels() As Byte, tmpSA As SafeArray1D, numBits As Long
        numBits = dstWidth * dstHeight * 4
        
        With tmpSA
            .cbElements = 1
            .cDims = 1
            .cLocks = 1
            .lBound = 0
            .cElements = numBits
            .pvData = m_tmpDIBBits
        End With
        
        PutMem4 VarPtrArray(imgPixels()), VarPtr(tmpSA)
        
        'The imgPixels() array now points at m_tmpDIBBits.
        
        'Iterate through all pixels and swap R/B channels.  This is much faster when compiled to
        ' native code with the "Remove Array Bounds Checks" option enabled.
        Dim x As Long, tmpColor As Byte
        For x = 0 To numBits - 1 Step 4
            tmpColor = imgPixels(x)
            imgPixels(x) = imgPixels(x + 2)
            imgPixels(x + 2) = tmpColor
        Next x
        
        'Finished!  Before exiting, we need to "unwrap" the imgPixels array from our GDI DIB.
        ' (Otherwise, VB will try to free that memory when this function exits - but because VB
        ' didn't allocate the memory, terrible things will happen.)
        PutMem4 VarPtrArray(imgPixels), 0&
        
        'Note the SVG handle, width, and height we used for this render.  If the next request is for the
        ' same SVG at the same size, we can reuse our intermediary GDI buffer as-is.
        m_lastSVG = srcImage.INTERNAL_GetSVGHandle()
        m_lastWidth = dstWidth
        m_lastHeight = dstHeight
        
    '/end "need to update GDI buffer"
    End If
    
    'The GDI buffer now contains the SVG at the desired size.  We now need to paint the GDI buffer
    ' onto the destination DC.  GDI's AlphaBlend function will do this for us.  See MSDN for details
    ' on the opacity value (it must be structured a specific way to request custom alpha-blending).
    Dim opacityFlags As Long
    opacityFlags = Int(svgOpacity * 2.55! + 0.5!) * &H10000 Or &H1000000
    AlphaBlend dstDC, dstX, dstY, dstWidth, dstHeight, m_tmpDC, 0, 0, dstWidth, dstHeight, opacityFlags
    
    'Report success
    INTERNAL_DrawTreeToDC = True
    
End Function

'Free a persistent SVG handle
Public Sub INTERNAL_FreeTree(ByVal srcTree As Long)
    If (srcTree <> 0) Then resvg_tree_destroy srcTree
    m_ActiveTree = m_ActiveTree - 1
End Sub

'Private support functions follow

'Ensure a GDI (DIB) surface exists at the supplied dimensions.
Private Sub EnsureGDISurfaceReady(ByVal srfWidth As Long, ByVal srfHeight As Long)
    
    'If we haven't created a surface yet, we obviously need to create one
    Dim newDIBNeeded As Boolean
    newDIBNeeded = (m_tmpDIB = 0)
    
    'If we already have a surface, see if its dimensions are OK
    If (Not newDIBNeeded) Then
        newDIBNeeded = (m_tmpDIBHeader.Width <> srfWidth) Or (m_tmpDIBHeader.Height <> srfHeight)
    End If
    
    'If we need a new DIB, create one now
    If newDIBNeeded Then
        
        'Free our current DIB, if one exists
        If (m_tmpDIB <> 0) Then
            If (m_tmpDC <> 0) Then SelectObject m_tmpDC, m_tmpOldObject
            DeleteObject m_tmpDIB
            m_tmpDIB = 0
        End If
        
        'Prepare the required header
        With m_tmpDIBHeader
            .Size = LenB(m_tmpDIBHeader)
            .Planes = 1
            .BitCount = 32
            .Width = srfWidth
            .Height = -srfHeight
            .ImageSize = srfWidth * srfHeight * 4
        End With
        
        'Create a compatible DC, as necessary (only one per session)
        If (m_tmpDC = 0) Then m_tmpDC = CreateCompatibleDC(0&)
        
        'Create the DIB and select it into the target DC
        m_tmpDIBBits = 0
        m_tmpDIB = CreateDIBSection(m_tmpDC, m_tmpDIBHeader, 0&, m_tmpDIBBits, 0&, 0&)
        If (m_tmpDIB <> 0) Then m_tmpOldObject = SelectObject(m_tmpDC, m_tmpDIB)
        
        'Ensure we have both a DIB handle, pointer, and DC.  If any of these are broken,
        ' this whole process must be abandoned.
        If (m_tmpDIB = 0) Or (m_tmpDIBBits = 0) Or (m_tmpDC = 0) Then
            
            If (m_tmpDC <> 0) Then
                If (m_tmpDIB <> 0) Then SelectObject m_tmpDC, m_tmpOldObject
                DeleteDC m_tmpDC
                m_tmpDC = 0
            End If
            
            If (m_tmpDIB <> 0) Then
                DeleteObject m_tmpDIB
                m_tmpDIB = 0
            End If
            
            m_tmpDIBBits = 0
            m_tmpDIBHeader.Width = 0
            m_tmpDIBHeader.Height = 0
            m_tmpDIBHeader.ImageSize = 0
            
            Exit Sub
            
        End If
        
    'If we can reuse our existing buffer as-is, we must still zero it out before proceeding.
    Else
        ZeroMemory m_tmpDIBBits, srfWidth * srfHeight * 4
    End If
    
    'The GDI surface is now ready for rendering.
    
End Sub

'Given a VB string, fill a byte array with matching UTF-8 data.
' RETURNS: TRUE if successful; FALSE otherwise.
Private Function UTF8FromString(ByRef strSource As String, ByRef dstUtf8() As Byte, Optional ByRef lenUTF8 As Long) As Boolean
    
    Const FUNC_NAME As String = "UTF8FromStrPtr"
    
    'Failsafe checks
    On Error GoTo UTF8FromStrPtrFail
    UTF8FromString = False
    
    Dim srcPtr As Long, srcLenInChars As Long
    srcPtr = StrPtr(strSource)
    srcLenInChars = Len(strSource)
    If (srcPtr = 0) Or (srcLenInChars = 0) Then Exit Function
    
    'Rely on default null-termination behavior for a slight perf boost
    srcLenInChars = -1
    
    'Use WideCharToMultiByte() to calculate the required size of the final UTF-8 array.
    Const CP_UTF8 As Long = 65001   'Fixed constant for UTF-8 "codepage" transformations
    lenUTF8 = WideCharToMultiByte(CP_UTF8, 0, srcPtr, srcLenInChars, 0, 0, 0, 0)
    
    'If the returned length is 0, WideCharToMultiByte failed.
    ' (This typically only happens when invalid character combinations are found.)
    If (lenUTF8 <= 0) Then
        InternalError FUNC_NAME, "WideCharToMultiByte did not return a valid buffer length (#" & Err.LastDllError & ", " & srcLenInChars & ")"
    
    'The returned length is non-zero.  Prep a buffer, then process the bytes.
    Else
        
        'Prep a buffer to receive the UTF-8 bytes.
        '
        'NOTE: Originally, the buffer boundary calculation used (-1) instead of (+1), as you'd expect
        ' given that we are explicitly declaring the string length instead of relying on null-termination.
        ' (So the function won't return a null-char, either.)  Unfortunately, this leads to unpredictable
        ' write access violations.  I'm not the first to encounter this (see http://www.delphigroups.info/2/fc/502394.html)
        ' but my Google-fu has yet to turn up an actual explanation for why this might occur.
        '
        'Avoiding the problem is simple enough - pad our output buffer with a few extra bytes, "just in case"
        ' WideCharToMultiByte misbehaves.
        Dim safeBufferBound As Long
        safeBufferBound = lenUTF8 + 1
        ReDim dstUtf8(0 To safeBufferBound) As Byte
        
        'Use the API to perform the actual conversion
        Dim finalResult As Long
        finalResult = WideCharToMultiByte(CP_UTF8, 0, srcPtr, srcLenInChars, VarPtr(dstUtf8(0)), lenUTF8, 0, 0)
        
        'Make sure the conversion was successful.  (There is generally no reason for it to succeed when
        ' calculating a buffer length, only to fail here, but better safe than sorry.)
        UTF8FromString = (finalResult <> 0)
        If (Not UTF8FromString) Then InternalError FUNC_NAME, "WideCharToMultiByte could not perform the conversion, despite returning a valid buffer length (#" & Err.LastDllError & ")."
        
    End If
    
    Exit Function
    
UTF8FromStrPtrFail:
    InternalError FUNC_NAME, "VB error: " & Err.Description & "(#" & Err.Number & ")"
End Function

Private Sub InternalError(ByRef funcName As String, ByRef errMsg As String)
    Debug.Print "WARNING! svgSupport." & funcName & "() error: " & errMsg
End Sub
