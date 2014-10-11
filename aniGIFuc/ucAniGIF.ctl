VERSION 5.00
Begin VB.UserControl ucAniGIF 
   BackStyle       =   0  'Transparent
   CanGetFocus     =   0   'False
   ClientHeight    =   1230
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   1305
   ClipBehavior    =   0  'None
   ClipControls    =   0   'False
   DrawStyle       =   2  'Dot
   ForeColor       =   &H80000006&
   HasDC           =   0   'False
   HitBehavior     =   0  'None
   MaskColor       =   &H80000014&
   PaletteMode     =   4  'None
   ScaleHeight     =   82
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   87
   ToolboxBitmap   =   "ucAniGIF.ctx":0000
   Windowless      =   -1  'True
End
Attribute VB_Name = "ucAniGIF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Acknowledgements:
' Uses Paul Caton's excellent ASM thunking routines to create timers as needed.
' -- His routines prevent needing to add VB timer controls and allows greater flexibility, no hWnds
' Uses Carles P.V.'s LZW compression logic to create stdPicture returned by AnimatedGIF property
' -- His routines allow the DIB to be converted to GIF before being converted to stdPicture

' Come back and visit every now & again to check for updates/bug fixes/enhancements:
' http://www.planetsourcecode.com/vb/scripts/ShowCode.asp?txtCodeId=68040&lngWId=1

' VB6 compatible only. VB5 users - ask a VB6 buddy to compile this for you
' This is a self-contained, image control designed specifically for GIFs, animated or not.
' The image control will not accept non-GIF images.

' Although this control can support overlapped animated GIFs, it is highly suggested that
' you do not. The result will be excessive, nearly continuous repainting. Every time the
' animated GIF lower in the Zorder changes, it triggers paint events for each overlapped control.
' Such controls would be our usercontrol, image controls, labels, lines, and shape controls.
' Controls below the animated GIF are also affected. When windowless controls get refreshed,
' VB asks each control below it in the zOrder to refresh the part of itself that is overlapped
' by the top level control currently being refreshed.

' Compile the control for best results/performance, ideally with all optimizations checked.
' When uncompiled, message boxes will make images disappear until message box is closed.

' The dotted border in design time does not disappear. It is purposely painted to:
'   - Identify the animated GIF control from a normal VB image control
'   - Show you the bounds of the overall image, all frames

' Lessons learned: Creating a windowless, DC-less animated usercontrol was rather painful.
'                  Developing the single DIB strip method was fun and challenging

' Incompatibility issues
'   1. Noted by Joe Jordan. When a windowless usercontrol is placed on a
' picturebox that has custom painting (i.e., gradients), you may get a black line at the
' top of the picturebox. Unfortunately, this appears to be a problem with the picturebox
' and not this usercontrol. There is a workaround.  After custom painting your picturebox,
' execute this line:  Me.MyPictureBox.Picture = Me.MyPictureBox.Image.
' By forcing the custom painting into its picture property, the black line disappears.
'   2. The LoadAnimatedGIF_Remote routine is not unicode compatible.
' It wraps a VB function that only supports ANSI

' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
'                                       CHANGE HISTORY
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =
' 23 Jan 09:
'   - Larry Riddle reported bug. When setting the Loops property during runtime, it would not be honored; fixed
'   - Restarting animation after stopped/paused would overwrite virtual memory unnecessarily; fixed
'   - Modified auto-sizing routine; should be more user-friendly. Added AutoSize property.
'       -- Control is automatically resized to fit scaled gif only in these cases:
'               AutoSize=True or Stretch=gfsActualSize or new GIF loaded during design time
'               When Stretch=gfsClipped, AutoSize is ignored
'   - Modified so it is compatible with MS Access which converts windowless to windowed
'       -- Also required adding stdPic property page for design-time AnimatedGIF property
'   - Changed SolidBkgColor property name to BackColor
'   - Changed SolidBkgUsed property name to BackStyle
'   - When all loops are completed, changed frame displayed to the final frame vs the 1st frame
'   - Several properties did not include the PropertyChanged statement; fixed
'   - Miscellaneous minor tweaks applied also
' 11 Nov 08:
'   - Added following properties: CurrentFrame, SolidBkgColor, SolidBkgColorUsed, SteppedDelay
'   - Changed FrameChanged to include additional boolean parameter: viaTimer
'   - Minor efficiency tweaks applied
'  7 Jul 07:
'   - reworked LoadAnimatedGIF_File to be more unicode robust
'       -- would fail to load unicode file name if path was also unicode
'  5 Jul 07:
'   - Simply added the FrameChanged public event
'   - Ensured c_ClrTblCopy was cleared everytime new GIF loaded
'  4 Jul 07:
'   - Added CacheColorTables and RestoreColorTables functions
' 27 Mar 07:
'   - Per request, added Pietro Cecchi suggestion to allow offsetting the GIF frame(s)
'       -- Added OffsetX and OffsetY properties. These properties allow tweaking the left/top rendered positions
'   - Added ability to retrieve & change palettes for the GIF
'       -- Set GetPalette & SetPalette functions
' 17 Mar 07:
'   - Tweaks for better efficiency, including:
'       - Reduced total bytes cached by this module
'       - When SetDIBColorTable called, now updates only number of palette entries needed vs always 256
'       - BuildDIBstrip routine now creates absolute smallest size DIB required
' 11 Mar 07:
'   - In some cases, 1st frame of transparent GIF may not draw transparent; fixed
' 10 Mar 07:
'   - Bug reported by Pietro Cecchi. Crash can occur when rendering; fixed
'       -- RenderFrame was checking c_BkBuff.hBmp for non-zero but should be checking .hDIB
'   - Changed Buffer DIB, when used, to 24bpp vs 32bpp
'   - Allowed Mirrored property to accept mirroring on both X & Y axis simultaneously
'       -- Modified MirrorGIF routine to accomodate the change
' 9 Mar 07.
'   - Completely revamped single DIB strip process
'       - Main strip remains the same; however
'       - If a buffer is needed, a separate, single frame, 32bpp DIB bitmap created
'       - If a buffer mask is needed, a separate, single frame, 1 bpp DIB bitmap created
'       -- Now able to get cpu usage down to zero or near zero in both IDE & compiled
'       -- Now able to mirror using stretchBlt; tons faster
'   - Added LoadAnimatedGIF_Remote to load GIF file remotely (i.e., URL, server UNC, local hard drive, etc)
'       - Added Public Event: RemoteLoadComplete to inform when remote GIF received
'       - Added Public Event: RemoteLoadFailure to inform when remote GIF reading failed
'   - When animation terminates due to loops being complete and animation restarted, the loop count was not reset
'   - Rewrote the MirrorGIF routine; added null image check in Mirrored property
'   - Rewrote the UpdateMask routine; extremely fast now
'   - Bug found by Pietro Cecchi: stretching produced bad results; flag not checked correctly in TransferFrame; fixed
' 7 Mar 07. My birthday, light update
'   - Found logic error in routine that compresses bitmap frame to GIF structure; could GPF during runtime; fixed
'   - Changed DelayAnimation property to allow displaying GIF w/o animation.
'       -- Requires user to set Me.Action=gfaPlay to start animation
' 6 Mar 07
'   - Tweaked for efficiency. Nearly 75% less CPU usage over previous version; almost no usage registered when compiled
'   - Added Mirrored property, allows you to flip (horizontally or vertically) the animation
'   - Enhanced LoadAnimatedGIF_Array and modified/tweaked most every routine
'   - Added a "Rendering Logic" RTF document to help visualize the animation process
' 4 Mar 07
'   - Moved timer termination into class termination. Potential of erasing memory prematurely & GPF on Win9x
'   - MouseUp events were being sent to user as MouseMove events. Thanx Soorya for pointing this out
'   - Added Enabled property; does not effect animation but does effect mouse events
'   - Added Refresh method; though you should never need to call this for windowless, DC-less controls
'   - Tweaked Stretch property to allow changing scale on the fly
' 3 Mar 07
'   - Added LoopsEnded public event
'   - Added LoopsRemaining property
'   - Changing animation actions were not effective
'   - ManageTimer was using wrong frame delay; type-o on my part
'   - ActualWidth/AcutalHeight now return values in user's scalemode
' 2 Mar 07: Initial version
' = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = = =

' Public Properties/Methods
' --------------------------
' Action - starts,pauses,stops animation
' ActualHeight - source height of the entire GIF
' ActualWidth - source width of the entire GIF
' AnimatedGIF - DesignTime Only, sets the GIF to use in the control (See Property for Runtime restriction)
' AutoSize - Forces control to size to GIF actual size
' BackColor - Optional solid color to render GIF over
' BackStyle - Determines if BackColor is ignored or not
' CacheColorTables - stores palette(s) for eventual call to RestoreColorTables. See below
' CurrentFrame - Returns the current frame rendered
' DelayAnimation - whether or not GIF will start by showing 1st frame now & then continue processing rest
' Enabled - whether or not control responds to mouse events
' FrameCount - Read Only. Number of frames within the GIF
' LoadAnimatedGIF_Array - method to assign a GIF during runtime (must be 1 or 2 dimensional array)
' LoadAnimatedGIF_File - method to assign a GIF during runtime
' LoadAnimatedGIF_Remote - method to assign a GIF during runtime from a URL or server
' Loops - determines how many loops to complete before animation terminates. Zero is infinite
' LoopsRemaining - Read Only. Number of loops remaining before animation terminates.
' MinFrameDelay - Minimal ms delay before next frame is displayed. Used when GIF frames encoded with zero milliseconds
' Mirrored - option to mirror animation either horizontally or vertically
' MouseIcon - option to set custom cursor when mouse is over the control
' MousePointer - variety of default cursors to use when mouse is over the control
' OffsetX - option to adjust left edge of rendered frame by +/- n pixels
' OffsetY - option to adjust top edge of rendered frame by +/- n pixels
' Refresh - allows control to be refreshed during runtime
' RestoreColorTables - restores the palette(s) cached by call to CacheColorTables
' SteppedDelay - Overrides individual frame delays to the value passed
' Stretch - enables various scaling options

' Public events....
' --------------------------
' add any additional events you think you may need
Public Event Click()
Attribute Click.VB_Description = "Occurs when the user presses and then releases a mouse button over an object"
Attribute Click.VB_MemberFlags = "200"
Public Event MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseDown.VB_Description = "Occurs when the user presses the mouse button while an object has the focus"
Public Event MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseUp.VB_Description = "Occurs when the user releases the mouse button while an object has the focus"
Public Event MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
Attribute MouseMove.VB_Description = "Occurs when the user moves the mouse"
Public Event LoopsEnded() ' fired when a looping gif terminates its loops naturally
Attribute LoopsEnded.VB_Description = "Event that occurs when all loops for the GIF have terminated"
Public Event FrameChanged(ByVal FrameIndex As Long, viaTimer As Boolean) ' fired each time a frame is rendered
Attribute FrameChanged.VB_Description = "Event that occurs when any frame is rendered manually or via a timer"
' note ^^ The viaTimer parameter is only True if active animation is running,
'         If changes/renders frame any other way (i.e., moving to next frame, resetting, etc), viaTimer=False

' By calling LoadAnimatedGIF_Remote, you will get one of these events
Public Event RemoteLoadComplete(ByVal gifWidth As Single, ByVal gifHeight As Single, ByRef Cancel As Boolean) ' notifies download of GIF from URL is complete
Attribute RemoteLoadComplete.VB_Description = "Event that occurs when a remotely loaded GIF has been validated"
Public Event RemoteLoadFailure() ' notifies failure to download GIF
Attribute RemoteLoadFailure.VB_Description = "Event that occurs when a remotely loaded GIF failed to be read"


'-Callback declarations for Paul Caton thunking magic----------------------------------------------
Private z_CbMem   As Long    'Callback allocated memory address
Private Declare Function GetModuleHandleA Lib "KERNEL32" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "KERNEL32" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function IsBadCodePtr Lib "KERNEL32" (ByVal lpfn As Long) As Long
Private Declare Function VirtualAlloc Lib "KERNEL32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFree Lib "KERNEL32" (ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Private Declare Sub RtlMoveMemory Lib "KERNEL32" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
'-------------------------------------------------------------------------------------------------

' used to create a stdPicture from a byte stream
Private Declare Function CreateStreamOnHGlobal Lib "ole32" (ByVal hGlobal As Long, ByVal fDeleteOnRelease As Long, ppstm As Any) As Long
Private Declare Function OleLoadPicture Lib "olepro32" (pStream As Any, ByVal lSize As Long, ByVal fRunmode As Long, riid As Any, ppvObj As Any) As Long
Private Declare Function GlobalAlloc Lib "KERNEL32" (ByVal uFlags As Long, ByVal dwBytes As Long) As Long
Private Declare Function GlobalLock Lib "KERNEL32" (ByVal hMem As Long) As Long
Private Declare Function GlobalUnlock Lib "KERNEL32" (ByVal hMem As Long) As Long
Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (ByRef Ptr() As Any) As Long

' Kernel32 APIs used
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Sub FillMemory Lib "kernel32.dll" Alias "RtlFillMemory" (ByRef Destination As Any, ByVal Length As Long, ByVal Fill As Byte)
' User32 APIs used
Private Declare Function SetTimer Lib "user32.dll" (ByVal hWnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32.dll" (ByVal hWnd As Long, ByVal nIDEvent As Long) As Long
Private Declare Function GetDC Lib "user32.dll" (ByVal hWnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hWnd As Long, ByVal hdc As Long) As Long
Private Declare Function SetRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function FillRect Lib "user32.dll" (ByVal hdc As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long
' GDI32 APIs used
Private Declare Function SetDIBColorTable Lib "gdi32.dll" (ByVal hdc As Long, ByVal un1 As Long, ByVal un2 As Long, ByRef pcRGBQuad As Long) As Long
Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function StretchBlt Lib "gdi32.dll" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function CreateDIBSection Lib "gdi32.dll" (ByVal hdc As Long, ByRef pBitmapInfo As BITMAPINFO, ByVal un As Long, ByRef lplpVoid As Long, ByVal handle As Long, ByVal dw As Long) As Long
Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function SetBrushOrgEx Lib "gdi32.dll" (ByVal hdc As Long, ByVal nXOrg As Long, ByVal nYOrg As Long, ByRef lppt As POINTAPI) As Long
Private Declare Function GetStretchBltMode Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32.dll" (ByVal hdc As Long, ByVal nStretchMode As Long) As Long
Private Declare Function Rectangle Lib "gdi32.dll" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function GetStockObject Lib "gdi32.dll" (ByVal nIndex As Long) As Long
Private Declare Function GetDIBits Lib "gdi32.dll" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, ByRef lpBits As Any, ByRef lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long

' APIs used to create files
Private Declare Function CloseHandle Lib "KERNEL32" (ByVal hObject As Long) As Long
Private Declare Function GetFileSize Lib "kernel32.dll" (ByVal hFile As Long, ByRef lpFileSizeHigh As Long) As Long
Private Declare Function ReadFile Lib "kernel32.dll" (ByVal hFile As Long, ByRef lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, ByRef lpNumberOfBytesRead As Long, ByRef lpOverlapped As Any) As Long
Private Declare Function CreateFileW Lib "KERNEL32" (ByVal lpFileName As Long, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function CreateFile Lib "KERNEL32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, ByVal lpSecurityAttributes As Long, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Private Declare Function GetFileAttributesW Lib "kernel32.dll" (ByVal lpFileName As Long) As Long
Private Declare Function GetFileAttributes Lib "kernel32.dll" Alias "GetFileAttributesA" (ByVal lpFileName As String) As Long
Private Declare Function SetFilePointer Lib "kernel32.dll" (ByVal hFile As Long, ByVal lDistanceToMove As Long, ByRef lpDistanceToMoveHigh As Long, ByVal dwMoveMethod As Long) As Long

' constants used
Private Const HALFTONE As Long = 4&  ' used for SetStretchBltMode API
Private Const PALETTECOUNT = 256&    ' number colors in 8 bit palette
Private Const BLACK_BRUSH As Long = 4&
Private Const WHITE_BRUSH As Long = 0&

' Standard Window UDTs
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Type POINTAPI
    X As Long
    Y As Long
End Type
Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type
Private Type SAFEARRAY2D        ' used as DMA overlay on a DIB
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    rgSABound(0 To 1) As SAFEARRAYBOUND
End Type
Private Type BITMAPINFOHEADER   ' used to create a DIB
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type
Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiPalette(0 To PALETTECOUNT - 1) As Long
End Type


' Custom UDTs
Private Type ColorTableSTRCUT
    Index As Long               ' 0=global; 1-xxx are local tables, if any
    Tables() As Long            ' 2D array (0 to 256, 0 to xxx); color table(s)
End Type
Private Type GIFcoreProperties
    Width As Long               ' overall GIF width
    Height As Long              ' overall height
    Loops As Long               ' Nr loops defined in GIF (0=infinite)
    ScaleCx As Single ' pre calculated ratios when scaling/stretching is used
    ScaleCy As Single ' pre calculated ratios when scaling/stretching is used
End Type
Private Type GIFframeProperties
    Dimensions As RECT          ' bounding rectangle of frame
    Delay As Long               ' length of time (ms) frame stays visible
    TblIndex As Long            ' the ColorTableSTRCUT this frame uses
    imgOffset As Long           ' file byte postion where image begins in gif file
                                ' & after decompressed, position where img begins in DIB strip
    TransIndex As Byte          ' which palette index is to be transparent
    IsTransparent As Byte       ' does frame use transparency (0=no, 1=yes)
    Disposal As Byte            ' the disposal method for this frame (0-7)
End Type
Private Type CoreDCInfo
    DC As Long                  ' off-screen DC
    hDib As Long                ' DIB created for the DC (our DIB strip)
    dibPtr As Long              ' pointer for DIB
    hBmp As Long                ' original monochrome bitmap when DC was created
End Type
Private Type BufferDCInfo
    DC As Long                  ' off-screen DC
    hDib As Long                ' DIB created for the DC
    dibPtr As Long              ' pointer for DIB
    hBmp As Long                ' original monochrome bitmap when DC was created
    hDibBW As Long              ' if monochrome buffer mask is needed, handle to the mask
    dibPtrBW As Long            ' pointer to the monochrome mask
End Type

Private Const gfsAutoSize As Long = 8 ' self-explanatory; appended to c_ScaleMode as needed
' Custom Enumerations
Public Enum ScaleGIFConstants    ' See Stretch property. Settings can force control to resize
    gfsClip = 0                  ' will never scale, nor stretch
    gfsScaleAlways = 1           ' will always scale up or down as needed
    gfsStretch = 2               ' will stretch to fit, not scaled, distorted
    gfsShrinkScaleToFit = 3      ' will only scale down if needed else scale is 1:1
    gfsActualSize = 4            ' self-explanatory
End Enum
Public Enum AnimationActions     ' See Action property
    gfaStop = 0                  ' stop on current frame, reset current frame to first
    gfaPlay = 1                  ' start/restart from current frame
    gfaPause = 2                 ' stop on current frame, do not reset frame nr
    gfaForward = 3               ' show next frame only, then pause
    gfaReset = 4                 ' restart from 1st frame
End Enum
Public Enum MirrorConstants     ' See MirrorGIF routine
    gfmNone = 0                 ' no mirroring
    gfmHorizontal = 1           ' mirror horizontally
    gfmVertical = 2             ' mirror vertically
    gfmHorAndVer = 3            ' mirror on X & Y axis
End Enum
Public Enum DelayModeConstants  ' See DelayAnimation property
    gfdNone = 0                 ' entire GIF is built before 1st frame is displayed
    gfdDelayStartup = 1         ' first frame is displayed immediately, then rest are processed
    gfdDoNotAnimate = 2         ' Same as gfdDelayStartup but does not initialize animation
End Enum
Public Enum GIFBackStyle
    gfbTransparent = 0&
    gfbSolid = 1&
End Enum
    
' All Class-Level variables are prefixed with c_
Private c_OffSetX As Long      ' Optional: modifies the left edge where rendering will occur on control
Private c_OffSetY As Long      ' Optional: modifies the top edge where rendering will occur

Private c_SolidBkgUsed As GIFBackStyle
Private c_MinDelay As Integer   ' any delay less than this value will use this value
Private c_AniLoops As Integer   ' during animation: loops remaining
Private c_aniState As AnimationActions ' current animation state
Private c_ScaleMode As ScaleGIFConstants ' image scaling options
Private c_curFrame As Long      ' during animation: which frame is being rendered
Private c_DelayLoad As DelayModeConstants     ' pauses processing during runtime until after first frame is rendered
Private c_Mirror As MirrorConstants ' optional horizontal and/or vertical mirroring
Private c_StepDelay As Long         ' optional overrride of per frame delay, 0 returns to actual delays
Private c_isWindowed As Boolean

Private c_DC As CoreDCInfo          ' holds the decompressed GIF, all frames
Private c_BkBuff As BufferDCInfo    ' holds the backbuffer & Mask if needed

Private c_gifProps As GIFcoreProperties ' overall GIF properties
Private c_Frames() As GIFframeProperties ' collection of individual frame properties
Private c_ColorTables As ColorTableSTRCUT  ' collection of color tables used in the GIF
Private c_maskTable() As Long       ' GIF Mask palette, B&W palette if needed

' following used during decompressing GIF & released immediately
Private c_DataLen() As POINTAPI     ' tracks frames file positions and sizes
Private c_aBuff() As Byte           ' general use byte array
Private c_DIBarray() As Byte        ' another general use array (used in ConvertStripToGIF & BuildDIBstrip)
Private c_aPOT() As Long            ' Power of 2 look up table
Private c_gifData() As Byte         ' source GIF data during design time only
Private c_ClrTblCopy As ColorTableSTRCUT ' back up of color table. See CacheColorTables, RestoreColorTables


' internal timer related variables
Private c_tmrOwner As Long
Private c_tmrID As Long
Private c_Ptr As Long               ' gif source data pointer during parsing; timer function pointer afterwards


' ////////////////////// PUBLIC PROPERTIES/METHODS \\\\\\\\\\\\\\\\\\\\\\

Public Property Let AnimatedGIF(aGif As StdPicture)
Attribute AnimatedGIF.VB_Description = "Returns/sets a graphic to be displayed in a control"
Attribute AnimatedGIF.VB_ProcData.VB_Invoke_PropertyPut = "StandardPicture"
    Set AnimatedGIF = aGif
End Property
Public Property Set AnimatedGIF(aGif As StdPicture)
    
    ' This can be called at any time to return the current frame as a stdPicture, but
    ' do not call this during runtime to assign a GIF to the control.
    ' The stdPicture you pass to this routine is most likely created by VB's
    ' LoadPicture function which will convert the GIF to a bitmap which is
    ' unusable by this control. Use LoadAnimatedGIF_File or LoadAnimatedGIF_Array instead.
    
    ' So why does this work, but LoadPicture doesn't?
    ' The reason, I believe, is that VB declares its IPicture interface differently
    ' when this is called vs when LoadPicture is called.  The stdPicture is an
    ' IPictureDisp interface which is a subinterface of the IPicture.  The IPicture
    ' has a .KeepOriginalFormat which is set to True here, but is set to False when
    ' LoadPicture is called. When that property is True, we can get the data we need
    ' You can try it if you like. Un-rem the following lines of code & see for yourself:
'        If Not aGif Is Nothing Then
'            Dim IPic1 As IPicture, IPic2 As IPicture, testPic As StdPicture
'            Set IPic1 = aGif
'            Set testPic=LoadPicture(..put a GIF path/filename here...)
'            Set IPic2 = testPic
'            Debug.Print "IPic1 KeepOriginal Format = "; IPic1.KeepOriginalFormat
'            Debug.Print "IPic2 KeepOriginal Format = "; IPic2.KeepOriginalFormat
'            Stop
'        End If
    
    If IsInUserMode = True Then Exit Property
    ' Not to be used during runtime
    
    Dim tBag As PropertyBag
    Dim sMode As ScaleGIFConstants
    
    If aGif Is Nothing Then
        Erase c_gifData()
    Else
        ' neat hack to get the entire GIF data from a stdPicture that loaded the GIF
        ' VB6 only, can create a NEW property bag object
        Set tBag = New PropertyBag
        tBag.WriteProperty "myGIF", aGif    ' 5 char name, don't change it
        ReDim c_gifData(1 To UBound(tBag.Contents) - 53) ' position where GIF data starts
        ' need to get the array this way, otherwise invalid data is returned
        Call ArrayFromVarRef(tBag.Contents, 54) ' 54 based on property name length
        Set tBag = Nothing
    End If
    Call LoadGif
    PropertyChanged "AnimatedGIF"
End Property
Public Property Get AnimatedGIF() As StdPicture
    If Not c_curFrame = 0& Then Set AnimatedGIF = ConvertStripToGIF(c_curFrame)
End Property

Public Property Get CurrentFrame() As Long
Attribute CurrentFrame.VB_Description = "Returns the current frame"
    ' returns the current frame; 0 indicates no gif loaded
    CurrentFrame = c_curFrame
End Property

Public Property Let AutoSize(ByVal newValue As Boolean)
Attribute AutoSize.VB_Description = "Resizes the control to the dimensions of the scaled image"
    ' If true, then resizes the control to fit the image
    ' Note: If Stretch=ActualSize then control will be resized also
    ' otherwise control is not resized when new images are applied
    ' Exception: During design-time, new images are always first
    '   displayed actual size, regardless of any scaling options
    If Not newValue = CBool(c_ScaleMode And gfsAutoSize) Then
        c_ScaleMode = c_ScaleMode Xor gfsAutoSize
        If newValue = True Then Call UserControl_Resize
        PropertyChanged "AutoSize"
    End If
End Property
Public Property Get AutoSize() As Boolean
    AutoSize = CBool(c_ScaleMode And gfsAutoSize)
End Property

Public Property Let Enabled(Enable As Boolean)
    ' Enables/disables the control. Disabled controls get no mouse/key/focus events
    ' IMPORTANT: To ensure this property tells VB to treat it as the
    ' control's enabled property, you must ensure the attributes are correct:
    ' IDE menu: Tools|Procedure Attributes
    ' Find Enabled in dropdown box, click "Advanced" button
    ' In "Procedure ID" dropdown, find & select Enabled, then click "Apply"
    UserControl.Enabled = Enable
    PropertyChanged "Enabled"
End Property
Public Property Get Enabled() As Boolean
Attribute Enabled.VB_Description = "Returns/sets a value that determines whether an object can respond to user-generated events"
Attribute Enabled.VB_UserMemId = -514
    Enabled = UserControl.Enabled
End Property
Public Property Let FrameCount(nrFrames As Long)
Attribute FrameCount.VB_Description = "Returns the number of frames within the GIF"
    ' dummy. Property is Read Only. Allows property to be displayed in property sheet
End Property
Public Property Get FrameCount() As Long
    If Not c_curFrame = 0& Then FrameCount = UBound(c_Frames)
End Property

Public Property Get ActualWidth() As Long
Attribute ActualWidth.VB_Description = "Original width of the overall GIF image"
    ' total size of entire GIF, in user's scalemode
    ActualWidth = Int(ScaleX(c_gifProps.Width, vbPixels, vbContainerSize))
End Property
Public Property Get ActualHeight() As Long
Attribute ActualHeight.VB_Description = "Original height of the overall GIF image"
    ' total size of entire GIF, in user's scalemode
    ActualHeight = Int(ScaleY(c_gifProps.Height, vbPixels, vbContainerSize))
End Property

Public Property Let Stretch(newScale As ScaleGIFConstants)
Attribute Stretch.VB_Description = "Returns/sets a value that determines whether a graphic resizes to fit the size of an Image control"
    If newScale < gfsClip Or newScale > gfsActualSize Then Exit Property
    If Not c_ScaleMode = (newScale And Not gfsAutoSize) Then
        Dim lAction As AnimationActions
        c_ScaleMode = (c_ScaleMode And gfsAutoSize) Or newScale
        lAction = c_aniState    ' cache
        Me.Action = gfaPause    ' pause if needed
        Call UserControl_Resize ' resize depending on c_ScaleMode
        Me.Action = lAction     ' restart animation if needed
        If Not lAction = gfaPlay Then UserControl.Refresh
        PropertyChanged "Stretch"
    End If
End Property
Public Property Get Stretch() As ScaleGIFConstants
    Stretch = (c_ScaleMode And Not gfsAutoSize)
End Property

Public Property Let MinFrameDelay(ByVal Delay As Long)
Attribute MinFrameDelay.VB_Description = "The minimum number of milliseconds a frame will remain before it is replaced."
    If Delay < 10 Then Delay = 10           ' ensure absolute minimum delay
    If Delay > 32700 Then Delay = 32700     ' ensure absolute maximum delay
    If Not Delay = c_MinDelay Then
        c_MinDelay = Delay
        PropertyChanged "Delay"
    End If
End Property
Public Property Get MinFrameDelay() As Long
    MinFrameDelay = c_MinDelay
End Property

Public Property Let Loops(Count As Long)
Attribute Loops.VB_Description = "The number of loops that will occur before animation stops. Zero will loop infintely"
    ' set Loops=0 for infinite looping
    ' to prevent animation on startup, set DelayAnimation=gfdDoNotAnimate
    ' to prevent further animation during runtime, set Action=gfaPause or Action=gfaStop
    If Not c_gifProps.Loops = Count Then
        c_gifProps.Loops = Abs(Count)
        c_AniLoops = c_gifProps.Loops
        PropertyChanged "Loops"
    End If
End Property
Public Property Get Loops() As Long
    Loops = c_gifProps.Loops
End Property

Public Property Get LoopsRemaining() As Long
Attribute LoopsRemaining.VB_Description = "Returns remaining number of loops before animation terminates"
    ' Returns zero if infinite looping
    If c_curFrame = 0& Then ' no gif loaded
        LoopsRemaining = 0&
    ElseIf c_gifProps.Loops = 0 Then ' infinite
        LoopsRemaining = &H7FFFFFFF ' a really high number
    Else
        LoopsRemaining = c_AniLoops ' loops remaining
    End If
End Sub

Public Property Let Mirrored(MirrorStyle As MirrorConstants)
Attribute Mirrored.VB_Description = "Mirrors images either horizontally or vertically"
    
    If MirrorStyle < gfmNone Or MirrorStyle > gfmHorAndVer Then Exit Property
    
    If c_curFrame = 0& Then
        c_Mirror = MirrorStyle
        PropertyChanged "Mirrored"
        
    ElseIf Not MirrorStyle = c_Mirror Then
        
        Dim lAction As Long
        lAction = c_aniState                ' cache current state
        Me.Action = gfaPause                ' pause animation
        MirrorGIF MirrorStyle, (IsInUserMode = False), IsInUserMode
        c_Mirror = MirrorStyle              ' save new value
        If lAction = gfaPlay Then           ' continue animation or refresh
            Me.Action = lAction
        Else
            UserControl.Refresh
        End If
        PropertyChanged "Mirrored"
    
    End If

End Property
Public Property Get Mirrored() As MirrorConstants
    Mirrored = c_Mirror
End Property

Public Property Let Action(Act As AnimationActions)
    ' start, stop, pause animation
    If c_curFrame = 0& Then Exit Property
    If IsInUserMode = False Then Exit Property
    Select Case Act
    Case gfaForward
        ManageTimers False, False
        c_curFrame = c_curFrame + 1 ' wrap back to 1st frame if needed
        If c_curFrame > UBound(c_Frames) Then c_curFrame = 1
        c_aniState = gfaPause
        UserControl.Refresh
        RaiseEvent FrameChanged(c_curFrame, False)
    Case gfaPause
        If Not c_aniState = gfaStop Then
            ManageTimers False, False
            c_aniState = gfaPause
        End If
    Case gfaPlay
        If c_aniState = gfaStop Then c_curFrame = 1&
        ManageTimers True, False
    Case gfaReset           ' basically a Stop & then Play call
        ManageTimers False, False
        c_AniLoops = c_gifProps.Loops ' reset loops remaining
        c_curFrame = 1      ' reset current frame
        UserControl.Refresh
        RaiseEvent FrameChanged(c_curFrame, False)
        ManageTimers True, False
    Case gfaStop
        ManageTimers False, False
        c_AniLoops = c_gifProps.Loops ' reset loops remaining
        c_curFrame = 1      ' reset current frame
        UserControl.Refresh
        RaiseEvent FrameChanged(c_curFrame, False)
    End Select
End Property
Public Property Get Action() As AnimationActions
Attribute Action.VB_Description = "Start, Stop and Pause animation"
Attribute Action.VB_MemberFlags = "400"
    Action = c_aniState
End Property

Public Property Let DelayAnimation(Mode As DelayModeConstants)
Attribute DelayAnimation.VB_Description = "During runtime, animation is paused until the first frame is loaded and displayed"
    ' property attempts to fast load the 1st frame of a gif
    ' then release process (like a DoEvents call)
    
    ' allow passing Mode OR'd, but the gfdDoNotAnimate takes precedence
    If Mode < gfdNone Or Mode > (gfdDelayStartup Or gfdDoNotAnimate) Then Exit Property
    If Not c_DelayLoad = Mode Then
        c_DelayLoad = Mode
        PropertyChanged "DelayAnimation"
    End If
End Property
Public Property Get DelayAnimation() As DelayModeConstants
    DelayAnimation = c_DelayLoad
End Property

Public Property Let MouseIcon(MousePic As StdPicture)
Attribute MouseIcon.VB_Description = "Sets a custom mouse icon"
    Set MouseIcon = MousePic
    PropertyChanged "MouseIcon"
End Property
Public Property Set MouseIcon(MousePic As StdPicture)
    ' same as VB's MouseIcon property
    On Error Resume Next
    Set UserControl.MouseIcon = MousePic
    If MousePic Is Nothing Then UserControl.MousePointer = vbDefault
    PropertyChanged "MouseIcon"
End Property
Public Property Get MouseIcon() As StdPicture
    Set MouseIcon = UserControl.MouseIcon
End Property

Public Property Let MousePointer(Pointer As MousePointerConstants)
Attribute MousePointer.VB_Description = "Returns/sets the type of mouse pointer displayed when over part of an object"
    ' same as VB's MousePointer property
    On Error Resume Next
    UserControl.MousePointer = Pointer
    PropertyChanged "MousePointer"
End Property
Public Property Get MousePointer() As MousePointerConstants
    MousePointer = UserControl.MousePointer
End Property

Public Property Let OffsetX(NewOffset As Long)
Attribute OffsetX.VB_Description = "Adjusts left edge of rendered frame by n pixels"
    ' option to adjust left edge used for rendering by +/- NewOffset pixels
    c_OffSetX = NewOffset
    PropertyChanged "OffsetX"
End Property
Public Property Get OffsetX() As Long
    OffsetX = c_OffSetX
    If IsInUserMode = False Or c_aniState <> gfaPlay Then UserControl.Refresh
End Property

Public Property Let OffsetY(NewOffset As Long)
Attribute OffsetY.VB_Description = "Adjusts top edge of rendered frame by n pixels"
    ' option to adjust top edge used for rendering by +/- NewOffset pixels
    c_OffSetY = NewOffset
    If IsInUserMode = False Or c_aniState <> gfaPlay Then UserControl.Refresh
    PropertyChanged "OffsetY"
End Property
Public Property Get OffsetY() As Long
    OffsetY = c_OffSetY
End Property

Public Property Get BackColor() As OLE_COLOR
Attribute BackColor.VB_Description = "Returns/sets the background color used to display graphics in an object"
    ' option to render frames over a solid backgroun color
    BackColor = UserControl.BackColor
End Property
Public Property Let BackColor(ByVal Color As OLE_COLOR)
    On Error Resume Next
    UserControl.BackColor = Color
    If Err Then
        Err.Clear
    Else
        If (c_SolidBkgUsed = gfbSolid Or c_isWindowed = True) And IsInUserMode = False Then UserControl.Refresh
        PropertyChanged "BackColor"
    End If
End Property
Public Property Get BackStyle() As GIFBackStyle
Attribute BackStyle.VB_Description = "Indicates whether the background is transparent or opaque"
    ' option to render frames over a solid backgroun color
    ' if property=False then SolidBkgColor property is ignored
    BackStyle = c_SolidBkgUsed
End Property
Public Property Let BackStyle(ByVal newStyle As GIFBackStyle)
    If newStyle < gfbTransparent Or newStyle > gfbSolid Then Exit Property
    c_SolidBkgUsed = newStyle
    If IsInUserMode = False And c_isWindowed = False Then UserControl.Refresh
    PropertyChanged "BackStyle"
End Property

Public Property Let SteppedDelay(ByVal newDelay As Long)
Attribute SteppedDelay.VB_Description = "Forces all frames to be rendered by the value provided. Zero erases SteppedDelay"
    ' property allows you to separate all frames by a specific delay
    ' Setting the property to zero, releases stepped values and the
    ' frame's individual delay time resumes as normal.
    
    ' Ideal for repositioning GIF between frames at a steady pace
    ' Note: FrameChanged event will fire with the viaTimer parameter=True
    If newDelay < 1& Then
        c_StepDelay = 0&
    Else
        c_StepDelay = newDelay
    End If
    PropertyChanged "SteppedDelay"
End Property
Public Property Get SteppedDelay() As Long
    SteppedDelay = c_StepDelay
End Property

Public Function LoadAnimatedGIF_File(ByVal FileName As String) As Boolean
Attribute LoadAnimatedGIF_File.VB_Description = "Optional method to assign a GIF"
    ' loads a GIF by file name during runtime
    Dim lRtn As Long, bUnicode As Boolean
    
    On Error Resume Next
    
    Const INVALID_HANDLE_VALUE = -1&
    Const GENERIC_READ As Long = &H80000000
    Const OPEN_EXISTING = &H3
    Const FILE_SHARE_READ = &H1
    Const FILE_ATTRIBUTE_ARCHIVE As Long = &H20
    Const FILE_ATTRIBUTE_HIDDEN As Long = &H2
    Const FILE_ATTRIBUTE_READONLY As Long = &H1
    Const FILE_ATTRIBUTE_SYSTEM As Long = &H4
    Const FILE_ATTRIBUTE_NORMAL = &H80&

    ' test to see if a file exists
    If isStringANSI(FileName) Then
        lRtn = GetFileAttributes(FileName)
    Else
        bUnicode = True
        lRtn = GetFileAttributesW(StrPtr(FileName))
    End If
    
    If Not lRtn = INVALID_HANDLE_VALUE Then
        
        Dim Flags As Long, Access As Long
        Dim Disposition As Long, Share As Long
        Dim hFile As Long
    
        Access = GENERIC_READ
        Share = FILE_SHARE_READ
        Disposition = OPEN_EXISTING
        Flags = FILE_ATTRIBUTE_ARCHIVE Or FILE_ATTRIBUTE_HIDDEN Or FILE_ATTRIBUTE_NORMAL _
                Or FILE_ATTRIBUTE_READONLY Or FILE_ATTRIBUTE_SYSTEM
    
        If bUnicode Then hFile = CreateFileW(StrPtr(FileName), Access, Share, 0&, Disposition, Flags, 0&)
        If hFile = 0 Then
            ' hFile should never be zero. It should be -1 (error) or a valid handle
            ' when hFile is zero, most likely API was called on a Win9x system
            ' so we will call the ANSI version and see if that returns a handle
            hFile = CreateFile(FileName, Access, Share, 0&, Disposition, Flags, 0&)
        End If
        If Not hFile = -1 Then
            ' get file size. VB's Len(file) fails if path is also unicode
            lRtn = GetFileSize(hFile, 0&)
            If lRtn Then
                SetFilePointer hFile, 0&, 0&, 0&
                ReDim c_gifData(1 To lRtn)
                ReadFile hFile, c_gifData(1), lRtn, Flags, ByVal 0&
                If Not Flags = lRtn Then Erase c_gifData()
            End If
            CloseHandle hFile
        End If
    End If
    LoadAnimatedGIF_File = CBool(LoadGif())
End Function

Public Sub LoadAnimatedGIF_Remote(ByVal RemotePath As String)
Attribute LoadAnimatedGIF_Remote.VB_Description = "Optional method to assign a GIF"
    ' loads a GIF by file name, URL, and/or UNC during runtime
    ' NOTE: This routine IS NOT unicode friendly
    If Len(RemotePath) = 0 Then Exit Sub
    On Error Resume Next
    UserControl.AsyncRead RemotePath, vbAsyncTypeByteArray, "remoteGIF", vbAsyncReadGetFromCacheIfNetFail
    If Err Then
        Err.Clear
        RaiseEvent RemoteLoadFailure
    End If
    
End Sub

Public Function LoadAnimatedGIF_Array(Bytes() As Byte) As Boolean
Attribute LoadAnimatedGIF_Array.VB_Description = "Optional method to assign a GIF"
    
    ' loads a GIF stream during runtime
    Dim lRtn As Long, SA As SAFEARRAY2D
    On Error GoTo EH
    
    ' test passed Byte() array for null array
    lRtn = IsArrayEmpty(VarPtrArray(Bytes))
    If Not lRtn = 0 Then
        ' get number of dimensions
        CopyMemory SA, ByVal lRtn, 2&
        If SA.cDims < 3 Then
            ' now get the entire safe array & size our array appropriately
            CopyMemory SA, ByVal lRtn, 16 + (SA.cDims * 8)
            If SA.cDims = 1 Then
                ReDim c_gifData(1 To (SA.cbElements * SA.rgSABound(0).cElements))
            Else
                ReDim c_gifData(1 To (SA.cbElements * (SA.rgSABound(0).cElements * SA.rgSABound(1).cElements)))
            End If
            ' transfer supplied bytes to our module-level array
            CopyMemory c_gifData(1), ByVal SA.pvData, UBound(c_gifData)
        End If
    End If
    LoadAnimatedGIF_Array = CBool(LoadGif())
EH:
End Function

Public Function GetPalette(ByVal FrameIndex As Long, paletteEntries() As Long) As Boolean
Attribute GetPalette.VB_Description = "Returns palette entries for a selected frame"
    ' Purpose: Return the palette used by the passed FrameIndex
    ' FrameIndex: if zero then the global palette is passed
    '   -- Otherwise the local palette, if used, will be passed
    ' paletteEntries will always be 4 bytes per pixel in BGR format
    
    ' Note: Each frame of a GIF can use either the global palette or a local palette but not both
    '   - Global: A general use palette used by any frame that does not have a local palette
    '   - Local: A palette designed just for that frame
    
    ' Function returns true if the paletteBytes() was filled, otherwise
    ' - When FrameIndex = 0 and function returns false, then no global palette exists,
    '   meaning that each frame uses a local palette
    ' - When FrameIndex > 0 and function returns false, then the frame uses the
    '   global palette and uses no local palette
    
    If c_curFrame = 0& Then Exit Function   ' no gif loaded
    
    Dim bRestart As Boolean
    If Not FrameIndex < 0 Then
    
        If c_aniState = gfaPlay Then
            Me.Action = gfaPause
            bRestart = True
        End If
    
        If FrameIndex = 0 Then
            ReDim paletteEntries(0 To c_ColorTables.Tables(PALETTECOUNT, 0) - 1)
            CopyMemory paletteEntries(0), c_ColorTables.Tables(0, 0), (UBound(paletteEntries) + 1) * 4&
            GetPalette = True
        ElseIf FrameIndex <= UBound(c_Frames) Then
            If Not c_Frames(FrameIndex).TblIndex = 0 Then
                ReDim paletteEntries(0 To c_ColorTables.Tables(PALETTECOUNT, c_Frames(FrameIndex).TblIndex) - 1)
                CopyMemory paletteEntries(0), c_ColorTables.Tables(0, c_Frames(FrameIndex).TblIndex), (UBound(paletteEntries) + 1) * 4&
                GetPalette = True
            End If
        End If
        If bRestart Then Me.Action = gfaPlay
    End If
End Function

Public Function SetPalette(ByVal FrameIndex As Long, paletteEntries() As Long) As Boolean
Attribute SetPalette.VB_Description = "Sets palette entries for a selected frame"
    
    ' Purpose: Set the palette used by the passed FrameIndex
    ' FrameIndex: if zero then the global palette is set
    '   -- Otherwise the local palette, if used, will be set
    ' paletteEntries must always be a 1D array & 4 bytes per pixel in BGR format
    
    ' Note: Each frame of a GIF can use either the global palette or a local palette but not both
    '   - Global: A general use palette used by any frame that does not have a local palette
    '   - Local: A palette designed just for that frame
    
    ' Function returns true if the GIF's palette was updated
    ' - When FrameIndex = 0 and function returns false, then no global palette exists to be udpated,
    '   meaning that each frame uses a local palette
    ' - When FrameIndex > 0 and function returns false, then the frame uses the
    '   global palette and does not use a local palette. No updates
    
    If c_curFrame = 0& Then Exit Function   ' no gif loaded
    If IsArrayEmpty(VarPtrArray(paletteEntries)) = 0& Then Exit Function
    
    Dim bRestart As Boolean
    If Not FrameIndex < 0 Then
        
        If c_aniState = gfaPlay Then
            Me.Action = gfaPause
            bRestart = True
        End If
        On Error Resume Next
        Dim UB As Long, LB As Long
        
        LB = LBound(paletteEntries)
        UB = UBound(paletteEntries)
        If Err Then
            Err.Clear
        ElseIf FrameIndex = 0 Then
            If Abs(UB - LB) + 1 >= c_ColorTables.Tables(PALETTECOUNT, 0) Then
                CopyMemory c_ColorTables.Tables(0, 0), paletteEntries(LB), c_ColorTables.Tables(PALETTECOUNT, 0) * 4&
                SetPalette = True
            End If
        ElseIf FrameIndex <= UBound(c_Frames) Then
            If Not c_Frames(FrameIndex).TblIndex = 0 Then
                If Abs(UB - LB) + 1 >= c_ColorTables.Tables(PALETTECOUNT, c_Frames(FrameIndex).TblIndex) Then
                    CopyMemory c_ColorTables.Tables(0, c_Frames(FrameIndex).TblIndex), paletteEntries(LB), c_ColorTables.Tables(PALETTECOUNT, c_Frames(FrameIndex).TblIndex) * 4&
                    SetPalette = True
                End If
            End If
        End If
        If bRestart Then
            Me.Action = gfaPlay
        Else
            UserControl.Refresh
        End If
    End If
    
End Function

Public Sub Refresh()
Attribute Refresh.VB_Description = "Forces a complete repaint of a object"
    UserControl.Refresh
End Sub

Public Function CacheColorTables() As Boolean
Attribute CacheColorTables.VB_Description = "Have control save GIF color tables in preparation of callig SetPalette and later restoring original palette"
    ' used to cache the original color table(s) of the GIF.
    ' The cached colors can be restored to the GIF via the
    ' RestoreColorTables function
    If c_ClrTblCopy.Index = 0 Then
        c_ClrTblCopy = c_ColorTables
        c_ClrTblCopy.Index = 1  ' flag indicating already cached
        CacheColorTables = True
    End If
End Function

Public Function RestoreColorTables(bReleaseCache As Boolean) As Boolean
Attribute RestoreColorTables.VB_Description = "Restores GIF color tables to those cached from CacheColorTables"
    ' if bReleaseCache then color table backup is destroyed
    If c_ClrTblCopy.Index Then
        c_ColorTables = c_ClrTblCopy
        c_ColorTables.Index = -1 ' forces next frame to validate color table
        If bReleaseCache Then
            Erase c_ClrTblCopy.Tables()
            c_ClrTblCopy.Index = 0  ' flag indicating no cache exists
        End If
        RestoreColorTables = True
    End If
End Function

' ////////////////////// USERCONTROL METHODS \\\\\\\\\\\\\\\\\\\\\\

Private Sub UserControl_AsyncReadComplete(AsyncProp As AsyncProperty)
    
    ' Event fired when LoadAnimatedGIF_Remote is attempting to retrieve a remotely located GIF
    
    Dim bLoad As Boolean
    If AsyncProp.BytesRead < 15 Then
    
        RaiseEvent RemoteLoadFailure
    
    Else
    
        Dim Cx As Long, Cy As Long
        
        ReDim c_gifData(1 To 10)
        ArrayFromVarRef AsyncProp.Value, LBound(AsyncProp.Value)
        
        Select Case Left$(LCase(StrConv(c_gifData, vbUnicode)), 6)
            Case "gif89a", "gif87a"
                CopyMemory Cx, c_gifData(7), 2& ' get gif's overall width
                CopyMemory Cy, c_gifData(9), 2& ' and height
                RaiseEvent RemoteLoadComplete(Int(ScaleX(Cx, vbPixels, vbContainerSize)), Int(ScaleY(Cy, vbPixels, vbContainerSize)), bLoad)
                bLoad = Not bLoad
                
            Case Else   ' not a valid GIF file
                RaiseEvent RemoteLoadFailure
        End Select
        Erase c_gifData()
    End If
    
    If bLoad = True Then
        ' get entire image now & prepare to process it
        ReDim c_gifData(1 To AsyncProp.BytesRead)
        ArrayFromVarRef AsyncProp.Value, LBound(AsyncProp.Value)
        Call LoadGif
    End If

End Sub

Private Sub UserControl_AsyncReadProgress(AsyncProp As AsyncProperty)

    ' Event fired when LoadAnimatedGIF_Remote is attempting to retrieve a remote location GIF
    
    If AsyncProp.StatusCode = vbAsyncStatusCodeError Then
        ' simply looking for an error. If so, we abort the Async Read
        CancelAsyncRead AsyncProp.PropertyName
        RaiseEvent RemoteLoadFailure
    End If

End Sub

Private Sub UserControl_Hide()
    Me.Action = gfaPause ' pause timer; control is no longer visible
End Sub

Private Sub UserControl_HitTest(X As Single, Y As Single, HitResult As Integer)
    HitResult = vbHitResultHit  ' allows selecting image during design time
End Sub

Private Sub UserControl_Click()
    RaiseEvent Click
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseDown(Button, Shift, Int(ScaleX(X, vbPixels, vbContainerPosition)), Int(ScaleY(Y, vbPixels, vbContainerPosition)))
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseMove(Button, Shift, Int(ScaleX(X, vbPixels, vbContainerPosition)), Int(ScaleY(Y, vbPixels, vbContainerPosition)))
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    RaiseEvent MouseUp(Button, Shift, Int(ScaleX(X, vbPixels, vbContainerPosition)), Int(ScaleY(Y, vbPixels, vbContainerPosition)))
End Sub

Private Sub UserControl_Paint()

    RenderFrame c_curFrame, UserControl.hdc
    If IsInUserMode = False Then    ' draw the dotted border
        Rectangle UserControl.hdc, 0, 0, UserControl.ScaleWidth, UserControl.ScaleHeight
    End If
End Sub

Private Sub UserControl_Initialize()
    UserControl.ScaleMode = vbPixels
    c_MinDelay = 50                 ' default settings for new controls
    c_ScaleMode = gfsScaleAlways    ' default scale option
    c_DelayLoad = gfdDelayStartup   ' default delay mode
    UserControl.DrawStyle = vbDot   ' this is the dotted border style
    UserControl.ForeColor = vbWindowFrame ' & color, in case they were changed by you
End Sub

Private Sub UserControl_Terminate()
    UnloadGIF
    zTerminate
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    With PropBag
        c_gifData = .ReadProperty("GIF", c_gifData())
        c_MinDelay = .ReadProperty("Delay", c_MinDelay)
        c_ScaleMode = .ReadProperty("Stretch", gfsScaleAlways)
        c_DelayLoad = .ReadProperty("DelayLoad", gfdDelayStartup)
        If c_DelayLoad = -1 Then c_DelayLoad = gfdDelayStartup ' patch for backward compatibility
        c_Mirror = .ReadProperty("Mirror", gfmNone)
        UserControl.Enabled = .ReadProperty("Enabled", True)
        Set UserControl.MouseIcon = .ReadProperty("MouseIcon", Nothing)
        UserControl.MousePointer = .ReadProperty("MousePointer", vbDefault)
        UserControl.BackColor = .ReadProperty("BkgFill", vbButtonFace)
        c_SolidBkgUsed = .ReadProperty("BkgUsed", gfbTransparent)
        c_StepDelay = .ReadProperty("StepDelay", 0&)
        c_OffSetX = .ReadProperty("OffsetX", 0&)
        c_OffSetY = .ReadProperty("OffsetY", 0&)
    End With
    If Not LoadGif() = 0& Then
        ' override loop count if user changed it in the property sheet
        c_gifProps.Loops = PropBag.ReadProperty("Loops", 0)
        c_AniLoops = c_gifProps.Loops
    End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    With PropBag
        .WriteProperty "GIF", c_gifData()
        .WriteProperty "Delay", c_MinDelay, 50
        .WriteProperty "Stretch", c_ScaleMode, gfsScaleAlways
        .WriteProperty "Loops", c_gifProps.Loops, 0
        .WriteProperty "DelayLoad", c_DelayLoad, gfdDelayStartup
        .WriteProperty "Mirror", c_Mirror, gfmNone
        .WriteProperty "MouseIcon", UserControl.MouseIcon, Nothing
        .WriteProperty "MousePointer", UserControl.MousePointer, vbDefault
        .WriteProperty "Enabled", UserControl.Enabled, True
        .WriteProperty "OffsetX", c_OffSetX, 0&
        .WriteProperty "OffsetY", c_OffSetY, 0&
        .WriteProperty "BkgFill", UserControl.BackColor, vbButtonFace
        .WriteProperty "BkgUsed", c_SolidBkgUsed, gfbTransparent
        .WriteProperty "StepDelay", c_StepDelay, 0&
    End With
End Sub

Private Sub UserControl_Resize()

    If Not c_curFrame = 0& Then
        Dim Cx As Long, Cy As Long, oCx As Long, oCy As Long
        
        If c_gifProps.Height < 0 Then Exit Sub ' prevents possible recursion while sizing
        
        oCx = UserControl.ScaleWidth
        oCy = UserControl.ScaleHeight
        ScaleToDestination Cx, Cy
        If oCx = Cx And oCy = Cy Then
            ' no change in size expected, simply refresh
            UserControl.Refresh
        Else
            c_gifProps.Height = -c_gifProps.Height ' flag to prevent potential recursion
            UserControl.Size ScaleX(Cx, vbPixels, vbTwips), ScaleY(Cy, vbPixels, vbTwips)
            ' in design view, resizing by the top left/right sizing handles, for some reason,
            ' does not honor the .Size call above. In this case, we will try to size the extender
            If Not ((Cx = UserControl.ScaleWidth) And (Cy = UserControl.ScaleHeight)) Then
                On Error Resume Next
                Extender.Width = ScaleX(Cx, vbPixels, vbContainerSize)
                Extender.Height = ScaleY(Cy, vbPixels, vbContainerSize)
                If Err Then Err.Clear
            End If
            c_gifProps.Height = -c_gifProps.Height
        End If
    End If
    
End Sub

Private Sub UserControl_Show()
    If UserControl.hWnd Then
        If c_isWindowed = False Then
            c_isWindowed = True
            UserControl.BackStyle = 1
        End If
    End If
    If c_gifProps.ScaleCx = 0! Then ScaleToDestination 0&, 0&
    If c_aniState = gfaPause Then ManageTimers True, False ' start timer if not already started
End Sub


' ////////////////////// SUPPORTING ROUTINES \\\\\\\\\\\\\\\\\\\\\\

Private Sub ArrayFromVarRef(inArray() As Byte, Offset As Long)
    ' helper function for Me.AnimatedGIF
    CopyMemory c_gifData(1), inArray(Offset), UBound(c_gifData)
End Sub

Private Function IsArrayEmpty(ByVal FarPointer As Long) As Long
  ' test to see if an array has been initialized & return its safearray pointer
  CopyMemory IsArrayEmpty, ByVal FarPointer, 4&
End Function

Private Function ByteAlignOnWord(ByVal BitDepth As Byte, ByVal Width As Long) As Long
    ' generic function to align byte range on dWord boundaries for any bit depth
    ByteAlignOnWord = (((Width * BitDepth) + &H1F) And Not &H1F&) \ &H8
End Function

Private Function LoadGif() As Long

    Dim newProps As GIFcoreProperties
    Dim nrItems As Long
    
    Call UnloadGIF        ' kill any timer, start with clean slate
    c_gifProps = newProps ' start a fresh UDT
    
    If Not IsArrayEmpty(VarPtrArray(c_gifData)) = 0& Then
    
        ' setup a "2^(0-8)" lookup table used by ParseGIF & BuildDIBstrip
        ReDim c_aPOT(0 To 8)
        c_aPOT(0) = 1
        For nrItems = 1 To 8
            c_aPOT(nrItems) = c_aPOT(nrItems - 1) * 2
        Next
        
        ' scan for key properties from the GIF file
        nrItems = ParseGIF()
        
        On Error GoTo EH
        If Not nrItems = 0 Then ' couldn't parse the file
            Erase c_aBuff()
            c_Ptr = 0
            ' transfer GIF to a DIB strip (all frames in one bitmap)
            If BuildDIBstrip(True) = False Then
                nrItems = 0
            Else
                If IsInUserMode = True Then
                    If nrItems = 1 Then RaiseEvent FrameChanged(nrItems, False)
                End If
            End If
        End If
    End If

EH:
    If nrItems = 0 Then
        If Err Then Err.Clear
        Erase c_gifData
        UnloadGIF
        UserControl.Refresh
    End If
    LoadGif = nrItems
    
End Function

Private Sub UnloadGIF()
    ' frees memory, releases any timer subclassing
    On Error Resume Next
    ManageTimers False, False
    Erase c_Frames()        ' clear all GIF frame info
    c_aniState = gfaStop    ' current animation state
    c_curFrame = 0&          ' reset frame index
    c_Ptr = 0&
    ' clean up GDI memory objects
    With c_DC
        If Not .hDib = 0& Then ' delete any DIB we created
            DeleteObject SelectObject(.DC, .hBmp)
            .hDib = 0&
            .hBmp = 0&
        End If
        If Not .DC = 0& Then   ' delete any DC
            DeleteDC .DC
            .DC = 0&
        End If
    End With
    With c_BkBuff
        If Not .hDib = 0& Then ' delete any DIB we created
            DeleteObject SelectObject(.DC, .hBmp)
            .hDib = 0&
            .hBmp = 0&
        End If
        If Not .hDibBW = 0& Then
            DeleteObject .hDibBW
            .hDibBW = 0&
        End If
        If Not .DC = 0& Then   ' delete any back buffer DC
            DeleteDC .DC
            .DC = 0&
        End If
    End With
    ' finally clear any used arrays & close GIF file if needed
    Erase c_ColorTables.Tables()    ' global & local palettes
    Erase c_ClrTblCopy.Tables()
    Erase c_maskTable()             ' mask palette
    Erase c_aPOT()                  ' power of two lookup (should already be cleared)
    Erase c_DataLen()               ' gif image positions within file (should already be cleared)
    c_ColorTables.Index = 0&        ' reset for next gif file
    c_ClrTblCopy.Index = 0&
    zTerminate
    If Err Then Err.Clear
End Sub

Private Function ParseGIF() As Long

    On Error Resume Next
    ' a modified routine from some of my other GIF postings
    ' This one is scaled back and skips many GIF blocks not needed for displaying.
    ' It also tracks a few file positions for later use in the BuildDIBstrip routine
    
    Dim gByte As Byte                       ' general purpose Byte
    Dim gLong As Long                       ' general purpose Long
    Dim gString As String                   ' general purpose String
    Dim lFrameCount As Integer              ' nr frames in GIF
    Dim bGlobalColorTable As Boolean        ' does a global exist; if not, then possibly corrupt gif
    Dim cFrame As GIFframeProperties        ' frame data
    Dim emptyFrame As GIFframeProperties    ' overall GIF data
    
    On Error GoTo ExitReadRoutine
    
    ReDim c_DataLen(1 To 10)  ' used to track start position of image description & size of compressed data
    c_Ptr = 1
    With c_gifProps
    
        ' read signature
        Call ReadGifFile_Variable(6)
        gString = LCase(StrConv(c_aBuff, vbUnicode))
        If Not (gString = "gif89a" Or gString = "gif87a") Then Exit Function
        
        ' skip to the global color table information
        c_Ptr = 11
        gByte = ReadGifFile_Byte()             ' the packed byte
        ' GIF Logical Screen Descriptor packed byte per specs
            ' bit pos = 0: nr of bits = 1 ' global color table used
            ' bit pos = 1: nr of bits = 3 ' original resolution
            ' bit pos = 4: nr of bits = 1 ' table sorted
            ' bit pos = 5: nr of bits = 3 ' palette bit depth (can be 0 thru 7 inclusive)
        
        c_Ptr = 14  ' move ahead to next byte after header
        If (gByte And 128) = 128 Then ' color table used? If so, read it
            c_ColorTables.Index = 0
            Call ReadGifFile_ColorTable(0, (gByte And &H7) + 1)
            bGlobalColorTable = True
        'Else no global color table; probably uses local color tables
        End If
    
    End With
    
    ' scan thru the entire file to find all the images & other key data
    With c_gifProps
    
        Do
            Select Case ReadGifFile_Byte()    ' read a single byte
            Case 0  ' block terminators
            Case 33 'Extension Introducer
                Select Case ReadGifFile_Byte() ' read the extension type
                
                Case 255    ' application extension, look for a loop count
                    ' the Netscape extension should always be before any images
                    ' Get the length of extension: will always be 11
                    gByte = ReadGifFile_Byte()
                    ' read next 11 bytes & see if it is a netscape extension
                    Call ReadGifFile_Variable(gByte)
                    gString = UCase(StrConv(c_aBuff, vbUnicode))
                    If gString = "NETSCAPE2.0" Then
                        ' within the data, we can extract the number of loops
                        ' an animated gif is suppose to run.
                        gByte = ReadGifFile_Byte()
                        If gByte = 3 Then   ' valid netscape extension
                            ' move ahead one byte & the next two is the loop count
                            c_Ptr = c_Ptr + 1
                            .Loops = (ReadGifFile_Integer And &HFFFF&) ' convert unsigned Integer to Long
                        Else
                            c_Ptr = c_Ptr - 1
                        End If
                    End If
                    SkipGifBlock
                
                Case 249    ' Graphic Control Label
                            ' (description of frame & is an optional block) 8 bytes
                    ' Graphic Control Extension (packed byte)
                        'bit pos = 0: nr of bits = 3 ' reserved
                        'bit pos = 3: nr of bits = 3 ' disposal method
                        'bit pos = 6: nr of bits = 1 ' user input flag
                        'bit pos = 7: nr of bits = 1 ' transparency
                    
                    cFrame = emptyFrame ' begin a new frame structure
                    With cFrame
                        .imgOffset = c_Ptr - 2 ' image starts here in the file
                        gByte = ReadGifFile_Byte() ' skip static byte (block size of 4)
                        ' get next byte which contains several properties
                        gByte = ReadGifFile_Byte()
                        
                        ' how is frame cleared after it is shown
                        .Disposal = ((gByte \ &H4) And &H3)
                        If .Disposal = 3 Then
                            If lFrameCount = 0& Then .Disposal = lFrameCount
                        ElseIf .Disposal > 3 Then
                            .Disposal = 0   ' treat reserved disposal codes as zero
                        End If
                        
                        ' how long does frame stay before being disposed & make 1/1000sec vs 1/100sec
                        .Delay = (ReadGifFile_Integer And &HFFFF&) * 10 ' convert unsigned Integer to Long
                        
                        .IsTransparent = (gByte And &H1) ' has transparency?
                        If .IsTransparent = 1 Then
                            .TransIndex = ReadGifFile_Byte  ' transparent index
                        Else
                            c_Ptr = .imgOffset + 7 ' skip transindex byte
                        End If
                        ' skip rest of block
                        Call SkipGifBlock
                    End With
                    
                Case Else   ' Comment block, plain text extension, or Unknown extension
                    Call SkipGifBlock
                End Select
                
            Case 44 ' Image Descriptor (image dimensions & color table)
                ' location of image within logical window
                ' GIF Image Descriptor (packed byte)
                    'bit pos = 0: nr of bits = 1 ' local color table used
                    'bit pos = 1: nr of bits = 1 ' interlaced
                    'bit pos = 2: nr of bits = 1 ' table sorted
                    'bit pos = 3: nr of bits = 2 ' reserved
                    'bit pos = 5: nr of bits = 3 ' palette bit depth
                With cFrame
                    ' mark position where image description starts
                    c_DataLen(lFrameCount + 1).Y = c_Ptr - 1
                    ' does image start here or did it start with Block 249 above?
                    If .imgOffset = 0 Then .imgOffset = c_DataLen(lFrameCount + 1).Y
                    
                    ' convert unsigned integers to long
                    .Dimensions.Left = (ReadGifFile_Integer And &HFFFF&)
                    .Dimensions.Top = (ReadGifFile_Integer And &HFFFF&)
                    .Dimensions.Right = (ReadGifFile_Integer And &HFFFF&)
                    .Dimensions.Bottom = (ReadGifFile_Integer And &HFFFF&)
                    
                    ' next byte indicates if local color table used
                    gByte = ReadGifFile_Byte()
                    If (gByte And 128) = 128 Then   ' local color table used?
                        .TblIndex = c_ColorTables.Index + 1 ' update ref & read table
                        Call ReadGifFile_ColorTable(.TblIndex, (gByte And &H7) + 1)
                        c_ColorTables.Index = .TblIndex     ' update table count
                    Else
                        If bGlobalColorTable = False Then Exit Function ' corrupted gif
                        ' frame says to use a global table, but no global color table exists
                    End If
                    ' skip the LZW byte & move to 1st byte of image
                    gByte = ReadGifFile_Byte()
                    SkipGifBlock ' done with block
                End With
                
                ' calculate image size in compressed bytes (includes local table if used)
                lFrameCount = lFrameCount + 1
                c_DataLen(lFrameCount).X = c_Ptr - c_DataLen(lFrameCount).Y
                If c_DataLen(lFrameCount).X < 3 Then ' then invalid image data
                    lFrameCount = lFrameCount - 1 ' roll back the frame
                Else
                    ReDim Preserve c_Frames(1 To lFrameCount)
                    If lFrameCount = UBound(c_DataLen) Then
                        ReDim Preserve c_DataLen(1 To lFrameCount + 5)
                    End If
                    c_Frames(lFrameCount) = cFrame
                    ' ensure logical window large enough for every frame
                    With c_Frames(lFrameCount).Dimensions
                        If .Left + .Right > c_gifProps.Width Then c_gifProps.Width = .Left + .Right
                        If .Top + .Bottom > c_gifProps.Height Then c_gifProps.Height = .Top + .Bottom
                    End With
                End If
                cFrame = emptyFrame ' always start with a new frame
                
            Case 59 ' Trailer (end of images)
                ' Although more images may exist, this flag tells us not to use any others
                Exit Do
            Case Else
                ' shouldn't happen; abort with what we have
                Exit Do
            End Select
        Loop
          
    End With
    ' got this far? good to go
          
ExitReadRoutine:
    If Err Then Err.Clear
    If Not lFrameCount = 0 Then
        If lFrameCount > UBound(c_Frames) Then ReDim Preserve c_Frames(1 To lFrameCount)
    End If
    ParseGIF = lFrameCount
    
End Function

'/==================================================================================
'   Read thru bytes until a zero-byte Block Terminator is found
'/==================================================================================
Private Sub SkipGifBlock()
    For c_Ptr = c_Ptr To UBound(c_gifData)
        If c_gifData(c_Ptr) = 0 Then Exit For
        c_Ptr = c_Ptr + c_gifData(c_Ptr)
    Next
    c_Ptr = c_Ptr + 1
End Sub
'/==================================================================================
'   Read one byte from an open file
'/==================================================================================
Private Function ReadGifFile_Byte() As Byte

    If c_Ptr > UBound(c_gifData) Then
        Err.Raise 53, "ReadGifFile", "End of File"
        Exit Function
    End If
    ReadGifFile_Byte = c_gifData(c_Ptr)
    c_Ptr = c_Ptr + 1
End Function
'/==================================================================================
'   Read an Integer (2 byte) value from an open file
'/==================================================================================
Private Function ReadGifFile_Integer() As Integer
    If c_Ptr + 1 > UBound(c_gifData) Then
        Err.Raise 53, "ReadGifFile", "End of File"
        Exit Function
    End If
    CopyMemory ReadGifFile_Integer, c_gifData(c_Ptr), 2&
    c_Ptr = c_Ptr + 2
End Function
'/==================================================================================
'   Read one or more bytes from the open gif file
'/==================================================================================
Private Sub ReadGifFile_Variable(ByVal nrBytes As Long)

    ReDim c_aBuff(0 To nrBytes - 1)
    If c_Ptr + nrBytes - 1 <= UBound(c_gifData) Then
        CopyMemory c_aBuff(0), c_gifData(c_Ptr), nrBytes
        c_Ptr = c_Ptr + nrBytes
    End If
EH:
End Sub
'/==================================================================================
'   Reads color tables inside GIF file and updates class collection
'/==================================================================================
Private Sub ReadGifFile_ColorTable(ByVal TableSlot As Long, ByVal BitDepth As Long)
    Dim c As Long
    ReDim Preserve c_ColorTables.Tables(0 To PALETTECOUNT, 0 To TableSlot)
    If c_Ptr + c_aPOT(BitDepth) * 3 > UBound(c_gifData) Then
        Err.Raise 53, "ReadGifFile", "End of File"
        Exit Sub
    End If
    For c = 0 To c_aPOT(BitDepth) - 1 ' convert RGB to BGR along the way
        c_ColorTables.Tables(c, TableSlot) = (c_gifData(c_Ptr) * &H10000) Or (c_gifData(c_Ptr + 1) * &H100&) Or c_gifData(c_Ptr + 2)
        c_Ptr = c_Ptr + 3
    Next
    ' cache number of palette entries used by this table
    c_ColorTables.Tables(PALETTECOUNT, TableSlot) = c_aPOT(BitDepth)
End Sub


Private Function BuildDIBstrip(InitialLoad As Boolean) As Boolean
    
    ' Function creates one DIB (a strip of frames as a single bitmap):
    
    ' The process I use creates a data stream (bytes) containing the GIF format
    ' for each frame of the GIF. That stream is sent to an API to convert stream to
    ' a stdPicture for each frame then render the stdPicture to the DIB strip.
    ' This is up to 3 times faster than decompressing the GIF frames by hand but has
    ' a drawback which is addressed. What's the drawback?
    ' Pixel palette indexes can change during stdPicture creation:
    
    ' ok; thru trial & error, I found that the stdPicture will rewrite the GIF
    ' palette indexes in certain cases. This happens when a palette color is repeated.
    ' In that case, the GIF will be re-written to use the 1st palette index that
    ' references that repeated color. When transparency is used, this prevents me
    ' from using masks on my DIB strip because the transparency index I parsed from
    ' the GIF file may no longer be valid if the indexes were re-indexed. So, a
    ' workaround is to ensure every palette index is unique. This, of course, is
    ' not an issue if each frame was stored as a stdPicture because the
    ' IPicture interface (stdPicture) must keep that info or create a mask for its use.
    ' But I want to use one GDI resource vs potentially dozens upon dozens of stdPictures.
    ' Ex: 100 frames, 32x32 GIF: what's better, 1 bmp or 100 stdPictures?
    ' The bytes used by both are about the same, but resources used are far more with stdPics.
    
    ' Note. With recent enhanchements, this routine modified so it can be called twice if
    ' user opts to delay animation on startup.
    
    Dim f As Long, aPtr As Long
    Dim frameStart As Long, frameStop As Long
    Dim uniquePal(0 To 767) As Byte ' used to ensure palette index uniqueness
    Dim stripBMP As BITMAPINFO
    Dim tPic As StdPicture
    
    On Error GoTo EH
    If InitialLoad Then   ' first time thru, we need to do some things
        Dim maxDataLen As Long, maxHeight As Long, totalWidth As Long
        Dim bTransparency As Byte, tDC As Long
        Dim bEraseAll As Boolean, bNeedMask As Boolean
        ' minimal image size optimizing used: assuming most GIF frames will be of similar
        ' height and width, calculate the total width of all frames.
        ' Our DIB strip will be horizontal
        bEraseAll = True
        For f = 1 To UBound(c_Frames)
            With c_Frames(f)
                ' add the widths of each frame - this will be our DIB's overall width
                ' and may be less than the GIF's overall width
                totalWidth = totalWidth + .Dimensions.Right
                ' keep track of largest frame height - this will be our DIB's overall height
                ' and may be less than the GIF's overall height
                If .Dimensions.Bottom > maxHeight Then maxHeight = .Dimensions.Bottom
                 ' transparency determines mask frame and mask color table creation
                bTransparency = bTransparency Or .IsTransparent
                ' guesstimate size of byte stream needed to create stdPic GIFs
                If c_DataLen(f).X > maxDataLen Then maxDataLen = c_DataLen(f).X
                Select Case .Disposal
                Case 0, 1
                    bEraseAll = False   ' may need a back buffer
                Case 2
                    bNeedMask = True    ' needs back buffer & mask unless all frames are code 2
                Case 3
                    bEraseAll = False   ' definitely needs back buffer
                    bNeedMask = True    ' definitely needs a mask
                End Select
            End With
        Next
        
        ' This routine can create up to 5 GDI objects, depending on GIF disposal codes:
        ' 1 8bpp DIB strip (containing all frames) and DC is always created
        ' 1 single frame 24bpp DIB and DC may be created if a back buffer is required
        ' 1 single frame monochrome DIB may be created if the buffer requires a mask
        
        ' when in IDE design, only need to worry about 1st frame
        If IsInUserMode = False Or f = 2 Then
            totalWidth = c_Frames(1).Dimensions.Right
            maxHeight = c_Frames(1).Dimensions.Bottom
            bNeedMask = False   ' never need a buffer or mask for single frames
        Else
            If bEraseAll = True Then
                bNeedMask = False   ' all frames are erased after drawn; no buffer or mask needed
            Else
                ' combination of disposal codes require a buffer and maybe a mask
                If bNeedMask = False Then
                    ' double check mask necessity
                    If c_Frames(1).IsTransparent = 1 Then
                        bNeedMask = True ' when 1st frame has transparency, mask is always needed
                    ElseIf bTransparency = 1 Then
                        ' 1st frame is not transparent, but if other frames are, then we
                        ' will need a mask if the 1st frame isn't same size of entire gif
                        If Not c_Frames(1).Dimensions.Bottom = c_gifProps.Height Then
                            bNeedMask = True
                        ElseIf c_Frames(1).Dimensions.Bottom < c_gifProps.Height Or c_Frames(1).Dimensions.Right < c_gifProps.Width Then
                            bNeedMask = True ' when 1st frame has transparency, mask is always needed
                        ElseIf Not c_Frames(1).Dimensions.Right = c_gifProps.Width Then
                            bNeedMask = True
                        End If
                    End If
                End If
                ' only possible error here would be lack of system resources
                If BuildBackBuffer(bNeedMask) = False Then Exit Function
            End If
        End If
        With stripBMP.bmiHeader            '  create the color dib strip
            .biBitCount = 8
            .biClrUsed = PALETTECOUNT
            .biHeight = maxHeight
            .biWidth = totalWidth
            .biPlanes = 1
            .biSize = 40
        End With
        With stripBMP                       ' fix to address problem discussed at top
            For f = 1 To PALETTECOUNT - 1&  ' ensure each palette entry is used only once
                uniquePal(f * 3& + 2&) = f  ' RGB (byte) - used by GIF stream
                .bmiPalette(f) = f          ' BGR (long) - used by DIB
            Next
        End With
        
        tDC = GetDC(0&)
        ' create a DC & select our newly created strip into it
        c_DC.DC = CreateCompatibleDC(tDC)
        If Not c_DC.DC = 0& Then
            c_DC.hDib = CreateDIBSection(tDC, stripBMP, 0&, c_DC.dibPtr, 0&, 0&) ' create the DIB
            If c_DC.hDib = 0& Then
                DeleteDC c_DC.DC
                c_DC.DC = 0&
            End If
        End If
        ReleaseDC 0&, tDC
        If (c_DC.DC = 0& Or c_DC.hDib = 0&) Then Exit Function
        c_DC.hBmp = SelectObject(c_DC.DC, c_DC.hDib)
        
        ' create an all-black mask array
        If bNeedMask = True Or bTransparency = 1 Then ReDim c_maskTable(0 To PALETTECOUNT - 1)
        
        ReDim c_DIBarray(1 To maxDataLen + 790) ' oversize array to handle any frame
        '^^ 790 = 768 for global palette;13 for header;8 for image control block;1 for EOF
        ' the array size IS EXACTLY correct when an 8-bit, transparent GIF file uses
        ' both global & local tables -- DO NOT reduce its size at all
        
        ' FYI: GIF header Layout?
        '  1-6   Signature (i.e., GIF87a or GIF89a)
        '  7-10  Logical Window - overall dimensions (width/height, 2 bytes each)
        '  11    Packed byte describing global color table
        '  12    Background Window color index
        '  13    Pixel aspect ratio
        '  14-xx Global color table if it exists
        CopyMemory c_DIBarray(1), c_gifData(1), 13&    ' copy logical window into stream
        CopyMemory c_DIBarray(14), uniquePal(0), PALETTECOUNT * 3& ' copy our unique palette into the global
        c_DIBarray(11) = c_DIBarray(11) Or 135        ' 135=our global exists & is 256 colors
        frameStart = 1&: frameStop = 1&
    
    Else    ' 2nd time thru
        
        If Not UBound(c_Frames) = 1& Then
            For f = 1& To PALETTECOUNT - 1& ' ensure each palette entry is used only once
                uniquePal(f * 3& + 2&) = f  ' RGB (byte) - used by GIF stream
                stripBMP.bmiPalette(f) = f  ' BGR (long) - used by DIB
            Next
            SetDIBColorTable c_DC.DC, 0, 256, stripBMP.bmiPalette(0)
        End If
        frameStart = 2&: frameStop = UBound(c_Frames)
        
    End If
    
    ' Populate our DIB strip... using stdPictures because it is up to 3 times faster than
    ' manually decompressing the GIF using VB algorithms alone. Just requires some
    ' extra work because of a quirk with how stdPicture can rewrite a GIF frame.
    
    c_ColorTables.Index = -1& ' will force next rendering in TransferFrame to update palette
    For f = frameStart To frameStop
        
        aPtr = 782&         ' next position in stream
        With c_Frames(f)
            ' the c_DataLen() array elements are set in the ParseGIF routine
            If Not c_DataLen(f).Y = .imgOffset Then
                ' get the image control block (always 8 bytes) & copy to stream
                ' FYI: GIF control block layout? (n/a for v87a)
                ' 1     Extension Introducer (fixed at 33)
                ' 2     Block ID (fixed at 249)
                ' 3     Remainng bytes in block after this byte (fixed at 4)
                ' 4     Transparency & other v89a options (packed byte)
                ' 5-6   Delay time in hundredths of seconds
                ' 7     Transparency Index
                ' 8     Block terminator (fixed at 0)
                CopyMemory c_DIBarray(aPtr), c_gifData(.imgOffset), 8&
                c_DIBarray(aPtr + 3&) = (c_DIBarray(aPtr + 3&) And Not 1)
                '^^ turn transparnecy off; otherwise, stdPicture will not write
                '   the correct index when it is rendered to my DIB strip
                aPtr = aPtr + 8&
            End If
            ' get the image description block (always 10 bytes), and the
            '   local color table (0 to 768 bytes) & compressed image (variable length)
            ' FYI: GIF image description block
            ' 1     Block ID (fixed at 44)
            ' 2-5   Frame left/top offsets, 2 bytes each
            ' 6-9   Frame width/height, 2 bytes each
            ' 10    Packed byte describing local color table
            ' 11-xx Local color table, if used
            ' xxxxx compressed image data where its 1st byte is the LZW compression size
            CopyMemory c_DIBarray(aPtr), c_gifData(c_DataLen(f).Y), c_DataLen(f).X
            ' ^^ includes EOI (end of image flag) & block terminator
            '       ensure logical window is at least as big as this frame (some GIFs
            '       are corrupted this way). First zero out Top/Left - not needed
            CopyMemory c_DIBarray(aPtr + 1), 0&, 4&
            ' now copy frame's width/height to logical window width/height
            CopyMemory c_DIBarray(7), c_DIBarray(aPtr + 5&), 4&
            If Not .TblIndex = 0& Then
                ' the local table starts 10 bytes after block starts & the byte before that
                ' one tells how many colors in the table; replace with our unique palette
                CopyMemory c_DIBarray(aPtr + 10&), uniquePal(0), _
                            c_aPOT((c_DIBarray(aPtr + 9&) And &H7) + 1&) * 3&
            End If
            aPtr = aPtr + c_DataLen(f).X    ' calc total bytes in this frame
            c_DIBarray(aPtr) = 59           ' add an EOF flag
            
            ' identify where in the strip this frame will begin
            If f = 1& Then .imgOffset = 0& Else _
                .imgOffset = c_Frames(f - 1&).imgOffset + c_Frames(f - 1&).Dimensions.Right
            
            ' Using stdPicture to decompress the GIF frame...
            ' let's create the stdPicture preserving all index values; actual colors
            ' are not important at this step, but palette indexes are very important.
            Set tPic = PictureFromByteStream(c_DIBarray(), aPtr)
            If Not tPic Is Nothing Then
                ' blt the stdPicture to our strip at the correct position
                tPic.Render c_DC.DC + 0&, .imgOffset + 0&, 0&, .Dimensions.Right + 0&, .Dimensions.Bottom + 0&, 0&, tPic.Height, tPic.Width, -tPic.Height, ByVal 0&
                ' ^^ FYI: the "+ 0&" above are required else errors
            End If
        End With
        
    Next
    
    If frameStart = 1& Then
        c_curFrame = 1&
        If IsInUserMode = True Then
            Call UserControl_Resize
            If Not (c_DelayLoad = gfdNone) Then
                If UBound(c_Frames) = 1& Then
                    InitialLoad = False ' single frame; no secondary recursion needed
                Else
                    MirrorGIF gfmNone, True, False
                    c_Mirror = c_Mirror Or &H20&
                    InitialLoad = (ManageTimers(True, True)) ' pause a few ms to let other processes continue
                End If
            Else
                BuildDIBstrip False ' continue building additional frames
            End If
        Else    ' done
            ' in design view, we will make control resize to actual size
            f = c_ScaleMode
            c_ScaleMode = gfsActualSize
            Call UserControl_Resize
            c_ScaleMode = f
            InitialLoad = False
        End If
    End If
    ' second time thru and/or finished processing remaining frames
    If InitialLoad = False Then ' clean up & animate as needed
        Erase c_DIBarray
        Erase c_aPOT
        Erase c_DataLen
        MirrorGIF gfmNone, (IsInUserMode = False), False
        If IsInUserMode Then
            Erase c_gifData
            If (c_DelayLoad And gfdDoNotAnimate) = gfdNone Then
                c_AniLoops = c_gifProps.Loops
                ManageTimers True, False  ' establish animation timer
            Else
                RaiseEvent FrameChanged(1, False)
            End If
        End If
    End If
    
    BuildDIBstrip = True
EH:
    If Err Then
'        Stop            ' troublehsooting only
'        Resume
        Err.Clear
    End If
End Function

Private Function RenderFrame(Index As Long, hdc As Long) As Boolean

 ' Function renders a frame as a result of the class timer firing or a container repaint
 
    If Index = 0 Then Exit Function
    ' ^^ visible control without image can get paint events; nothing to draw
 
    Dim drawRect As RECT                   ' frame's bounding rect
    Dim bStretch As Boolean
    Dim d3Mask() As Byte, d3Color() As Byte ' used to cache current buffer contents as needed
    Dim hBrush As Long
    
    Const Ratio1to1 As Single = 1
    
    ' definitively determine if image will be stretched.
    ' Stretching requires extra lines of code to be executed in TransferFrame including modifying the target DC
    ' Just cause the Stretch property is set to gfsStretch we may not actually be stretching
    If Not c_gifProps.ScaleCx = Ratio1to1 Then
        bStretch = True
    ElseIf Not c_gifProps.ScaleCy = Ratio1to1 Then
        bStretch = True
    End If
    
    If c_BkBuff.hDib = 0& Then  ' no offscreen buffer, no mask
       ' this is the easy render method
       With c_Frames(Index).Dimensions
           SetRect drawRect, (.Left + c_OffSetX) * c_gifProps.ScaleCx, (.Top + c_OffSetY) * c_gifProps.ScaleCy, _
               .Right * c_gifProps.ScaleCx, .Bottom * c_gifProps.ScaleCy
       End With
       TransferFrame Index, hdc, bStretch, drawRect
       
    Else
        ' append to the running mask, this frame's mask
        UpdateMask Index, False, d3Mask(), d3Color()
        
        ' copy current frame to offscreen buffer
        drawRect = c_Frames(Index).Dimensions
        TransferFrame Index, c_BkBuff.DC, False, drawRect
        
        ' now transfer offscreen to screen
        SetRect drawRect, c_OffSetX, c_OffSetY, UserControl.ScaleWidth, UserControl.ScaleHeight
        TransferFrame -Index, hdc, bStretch, drawRect
        
        ' erase or update the running mask
        If c_Frames(Index).Disposal > 1 Then UpdateMask Index, True, d3Mask(), d3Color()
        
     End If
    RenderFrame = True

End Function

Private Sub TransferFrame(ByVal frameNr As Long, hdc As Long, _
                bStretch As Boolean, destR As RECT)

    ' function BLTs the gif frame onto the passed DC while scaling
    ' It also creates a mask on the fly using a niffty color table tweak.
    ' Majority of code is keeping track of color tables during the task & setting offsets
    
    ' when clipping, a frame can be completely outside of the DC; simply abort
    If destR.Bottom < 1 Or destR.Right < 1 Then Exit Sub
    
    Dim mROP As Long, dcRect As RECT, hBrush As Long
    Dim oldBrEx As POINTAPI, oldStretchMode As Long
    Dim xOffset As Long, xWidth As Long, xHeight As Long
    Dim srcDC As Long, lTransColor As Long
        
    With c_Frames(Abs(frameNr))
    
        If bStretch = True Then ' set target DC stretch mode & cache original setting
            oldStretchMode = SetStretchBltMode(hdc, HALFTONE)
            SetBrushOrgEx hdc, 0&, 0&, oldBrEx
        End If
        
        If (c_SolidBkgUsed = gfbSolid Or c_isWindowed = True) Then
            SetRect dcRect, 0, 0, c_gifProps.Width * c_gifProps.ScaleCx, c_gifProps.Height * c_gifProps.ScaleCy
            If UserControl.BackColor < 0& Then
                hBrush = CreateSolidBrush(GetSysColor(UserControl.BackColor And &HFF))
            Else
                hBrush = CreateSolidBrush(UserControl.BackColor)
            End If
            FillRect hdc, dcRect, hBrush
            DeleteObject hBrush
        End If
        
        
        If frameNr < 0& Then
            ' render to screen from the mask frame & back buffer
            xWidth = c_gifProps.Width
            xHeight = c_gifProps.Height
            
            SelectObject c_BkBuff.DC, c_BkBuff.hDibBW
            If bStretch = True Then
                StretchBlt hdc, destR.Left, destR.Top, destR.Right, destR.Bottom, _
                    c_BkBuff.DC, 0&, 0&, xWidth, xHeight, vbSrcAnd
            Else    ' clipping only
                BitBlt hdc, destR.Left, destR.Top, destR.Right, destR.Bottom, c_BkBuff.DC, 0&, 0&, vbSrcAnd
            End If
            SelectObject c_BkBuff.DC, c_BkBuff.hDib
            mROP = vbSrcPaint           ' set ROP for the color portion
            srcDC = c_BkBuff.DC ' the back buffer will now be used to render color to screen
        
        Else    ' otherwise we are transfering frame to passed DC (destination DC)
            
            srcDC = c_DC.DC
            xOffset = .imgOffset
            xWidth = .Dimensions.Right
            xHeight = .Dimensions.Bottom
            
            If .IsTransparent = 1 Then ' color table hack to create masks for paletted images
            
                ' The c_maskTable is all black, no other colors
                c_maskTable(.TransIndex) = vbWhite   ' set transparent index to white
                SetDIBColorTable c_DC.DC, 0, c_ColorTables.Tables(PALETTECOUNT, .TblIndex), c_maskTable(0) ' put the table to the DIB
                c_maskTable(.TransIndex) = vbBlack ' set the color back to black
                mROP = vbSrcPaint       ' set ROP for the color portion
                ' now draw the mask to the destination DC, note the ROP used >>
                If bStretch = True Then
                    StretchBlt hdc, destR.Left, destR.Top, destR.Right, destR.Bottom, _
                        srcDC, xOffset, 0&, xWidth, xHeight, vbSrcAnd
                Else    ' clipping only
                    BitBlt hdc, destR.Left, destR.Top, destR.Right, destR.Bottom, srcDC, xOffset, 0&, vbSrcAnd
                End If
                
            Else
                mROP = vbSrcCopy        ' no transparency used; default ROP
            End If
    
            ' do we need to select a different table for the current frame?
            If .IsTransparent = 1 Or Not .TblIndex = c_ColorTables.Index Then
                If .IsTransparent = 1 Then
                    ' cache original transparency color & change it to black
                    lTransColor = c_ColorTables.Tables(.TransIndex, .TblIndex)
                    c_ColorTables.Tables(.TransIndex, .TblIndex) = vbBlack
                End If
                c_ColorTables.Index = .TblIndex  ' update current table ref
                ' change the dib colors & replace transparency color
                SetDIBColorTable c_DC.DC, 0, c_ColorTables.Tables(PALETTECOUNT, .TblIndex), c_ColorTables.Tables(0, .TblIndex)
                If .IsTransparent = 1 Then c_ColorTables.Tables(.TransIndex, .TblIndex) = lTransColor
            End If
            
        End If
        ' blt the color image and done
        If bStretch = True Then
            StretchBlt hdc, destR.Left, destR.Top, destR.Right, destR.Bottom, _
                srcDC, xOffset, 0&, xWidth, xHeight, mROP
                ' reset to user DC's original settings
                SetStretchBltMode hdc, oldStretchMode
                SetBrushOrgEx hdc, oldBrEx.X, oldBrEx.Y, oldBrEx
        Else
            BitBlt hdc, destR.Left, destR.Top, destR.Right, destR.Bottom, srcDC, xOffset, 0&, mROP
        End If
        
    End With

End Sub

Private Sub UpdateMask(Index As Long, bDisposing As Boolean, d3Mask() As Byte, d3Color() As Byte)

    ' routine maintains a running mask for the GIF

    Dim Rows As Long, gD3row As Long
    Dim mOffset As Long, mScanWidth As Long
    Dim maskBytes() As Byte, colorBytes() As Byte
    Dim maskSA As SAFEARRAY2D, colorSA As SAFEARRAY2D
    Dim eRect As RECT, hBrush As Long

    If Index = 1 Or c_Frames(Index).Disposal = 3 Then
        If Not c_BkBuff.hDibBW = 0& Then
            With maskSA                 ' DIB strip overlay
                .cbElements = 1
                .cDims = 2
                .pvData = c_BkBuff.dibPtrBW
                .rgSABound(0).cElements = c_gifProps.Height
                .rgSABound(1).cElements = ByteAlignOnWord(1, c_gifProps.Width)
            End With
            CopyMemory ByVal VarPtrArray(maskBytes), VarPtr(maskSA), 4&
        End If
        
        With colorSA                ' Buffer overlay
            .cbElements = 1
            .cDims = 2
            .pvData = c_BkBuff.dibPtr
            .rgSABound(0).cElements = c_gifProps.Height
            .rgSABound(1).cElements = ByteAlignOnWord(24, c_gifProps.Width)
        End With
        CopyMemory ByVal VarPtrArray(colorBytes), VarPtr(colorSA), 4&
    End If
    
    If bDisposing Then
        
        With c_Frames(Index)
            Select Case .Disposal
            
            Case 2 ' erase area just drawn
                SetRect eRect, .Dimensions.Left, .Dimensions.Top, .Dimensions.Left + .Dimensions.Right, .Dimensions.Top + .Dimensions.Bottom
                FillRect c_BkBuff.DC, eRect, GetStockObject(BLACK_BRUSH)
                If Not c_BkBuff.hDibBW = 0& Then
                    SelectObject c_BkBuff.DC, c_BkBuff.hDibBW
                    FillRect c_BkBuff.DC, eRect, GetStockObject(WHITE_BRUSH)
                    SelectObject c_BkBuff.DC, c_BkBuff.hDib
                End If
                
            Case 3 ' replace with previous contents
                mOffset = .Dimensions.Left * 3&
                mScanWidth = .Dimensions.Right * 3&
                For Rows = c_gifProps.Height - .Dimensions.Top - 1& To c_gifProps.Height - .Dimensions.Bottom - .Dimensions.Top Step -1&
                    CopyMemory colorBytes(mOffset, Rows), d3Color(0, gD3row), mScanWidth
                    gD3row = gD3row + 1&
                Next
                If Not c_BkBuff.hDibBW = 0& Then
                    gD3row = 0&
                    mOffset = .Dimensions.Left \ 8
                    mScanWidth = ByteAlignOnWord(1, .Dimensions.Right)
                    For Rows = c_gifProps.Height - .Dimensions.Top - 1& To c_gifProps.Height - .Dimensions.Bottom - .Dimensions.Top Step -1&
                        CopyMemory maskBytes(mOffset, Rows), d3Mask(0, gD3row), mScanWidth
                        gD3row = gD3row + 1&
                    Next
                End If
            End Select
        End With
            
    Else            ' copy new frame onto existing
        
        With c_Frames(Index)
            If Index = 1 Then   ' first frame, erase
                ' fill buffer with black (anti-mask)
                FillMemory colorBytes(0, 0), colorSA.rgSABound(1).cElements * c_gifProps.Height, 0
                ' fill mask with white
                If Not c_BkBuff.hDibBW = 0& Then FillMemory maskBytes(0, 0), maskSA.rgSABound(1).cElements * c_gifProps.Height, 255
            End If
            
            If .Disposal = 3 Then   ' cache current buffer and mask contents as needed
            
                mOffset = .Dimensions.Left * 3&
                mScanWidth = .Dimensions.Right * 3&
                ReDim d3Color(0 To mScanWidth - 1&, 0 To .Dimensions.Bottom - 1&)
                For Rows = c_gifProps.Height - .Dimensions.Top - 1& To c_gifProps.Height - .Dimensions.Bottom - .Dimensions.Top Step -1&
                    CopyMemory d3Color(0, gD3row), colorBytes(mOffset, Rows), mScanWidth
                    gD3row = gD3row + 1&
                Next
                If Not c_BkBuff.hDibBW = 0& Then
                    gD3row = 0&
                    mScanWidth = ByteAlignOnWord(1, .Dimensions.Right)
                    mOffset = .Dimensions.Left \ 8
                    ReDim d3Mask(0 To mScanWidth - 1&, 0 To .Dimensions.Bottom - 1&)
                    For Rows = c_gifProps.Height - .Dimensions.Top - 1& To c_gifProps.Height - .Dimensions.Bottom - .Dimensions.Top Step -1&
                        CopyMemory d3Mask(0, gD3row), maskBytes(mOffset, Rows), mScanWidth
                        gD3row = gD3row + 1&
                    Next
                End If
                
            End If
            
        End With
        
        If Not c_BkBuff.hDibBW = 0& Then    ' update mask as needed
        
            SelectObject c_BkBuff.DC, c_BkBuff.hDibBW
            With c_Frames(Index)
            
                ' transfer frame's mask to the buffer's mask
            
                If .IsTransparent = 1 Then
                    c_maskTable(.TransIndex) = vbWhite
                    SetDIBColorTable c_DC.DC, 0&, c_ColorTables.Tables(PALETTECOUNT, .TblIndex), c_maskTable(0) ' put the table to the DIB
                    c_maskTable(.TransIndex) = vbBlack
                    BitBlt c_BkBuff.DC, .Dimensions.Left, .Dimensions.Top, .Dimensions.Right, .Dimensions.Bottom, c_DC.DC, .imgOffset, 0&, vbSrcAnd
                Else
                    SetDIBColorTable c_DC.DC, 0&, c_ColorTables.Tables(PALETTECOUNT, .TblIndex), c_maskTable(0) ' put the table to the DIB
                    BitBlt c_BkBuff.DC, .Dimensions.Left, .Dimensions.Top, .Dimensions.Right, .Dimensions.Bottom, c_DC.DC, .imgOffset, 0&, vbSrcCopy
                End If
                
            End With
            
            SelectObject c_BkBuff.DC, c_BkBuff.hDib
            c_ColorTables.Index = -1&    ' forces color frames to reselect its palette before rendering it
        
        End If
    
    End If
    
    If Not maskSA.pvData = 0& Then CopyMemory ByVal VarPtrArray(maskBytes), 0&, 4&
    If Not colorSA.pvData = 0& Then CopyMemory ByVal VarPtrArray(colorBytes), 0&, 4&
    
End Sub

Private Sub MirrorGIF(newState As MirrorConstants, SingleFrame As Boolean, MirrorMask As Boolean)

    Dim Index As Long, mirrorState As MirrorConstants
    Dim frameNr As Long, nrFrames As Long
    Dim mirrorX As Long, mirrorY As Long
    Dim mirrorCx As Long, mirrorCy As Long
    
    If (c_Mirror And &H20) = &H20 Then ' &H20 is flag indicating we been thru here before
        frameNr = 2& ' routine called second time
        c_Mirror = c_Mirror And Not &H20
    Else
        frameNr = 1& ' routine called first time
    End If
    
    ' remove any mirror options that are already applied
    mirrorState = c_Mirror Xor newState
    If mirrorState = gfmNone Then Exit Sub ' nothing to do
    
    If SingleFrame Then
        nrFrames = 1&    ' delay start up or IDE design, always one frame
    Else
        nrFrames = UBound(c_Frames)
    End If
    
    For Index = frameNr To nrFrames
        ' flip the color frames
        With c_Frames(Index)
            If (mirrorState And gfmHorizontal) = gfmHorizontal Then
                mirrorX = .imgOffset + .Dimensions.Right - 1&
                mirrorCx = -.Dimensions.Right
                .Dimensions.Left = c_gifProps.Width - (.Dimensions.Right + .Dimensions.Left)
            Else
                mirrorX = .imgOffset
                mirrorCx = .Dimensions.Right
            End If
            If (mirrorState And gfmVertical) = gfmVertical Then
                mirrorY = .Dimensions.Bottom - 1&
                mirrorCy = -.Dimensions.Bottom
                .Dimensions.Top = c_gifProps.Height - (.Dimensions.Bottom + .Dimensions.Top)
            Else
                mirrorY = 0&
                mirrorCy = .Dimensions.Bottom
            End If
            ' flip each frame
            StretchBlt c_DC.DC, .imgOffset, 0&, .Dimensions.Right, .Dimensions.Bottom, _
                c_DC.DC, mirrorX, mirrorY, mirrorCx, mirrorCy, vbSrcCopy
        End With
    Next
    If MirrorMask Then
        If Not c_BkBuff.hDib Then
            If (mirrorState And gfmHorizontal) = gfmHorizontal Then
                mirrorX = c_gifProps.Width - 1&
                mirrorCx = -c_gifProps.Width
            Else
                mirrorX = 0&
                mirrorCx = c_gifProps.Width
            End If
            If (mirrorState And gfmVertical) = gfmVertical Then
                mirrorY = c_gifProps.Height - 1&
                mirrorCy = -c_gifProps.Height
            Else
                mirrorY = 0&
                mirrorCy = c_gifProps.Height
            End If
            ' flip the buffer
            StretchBlt c_BkBuff.DC, 0&, 0&, c_gifProps.Width, c_gifProps.Height, _
                c_BkBuff.DC, mirrorX, mirrorY, mirrorCx, mirrorCy, vbSrcCopy
        
            If Not c_BkBuff.hDibBW = 0& Then
                ' flip the mask
                SelectObject c_BkBuff.DC, c_BkBuff.hDibBW
                StretchBlt c_BkBuff.DC, 0&, 0&, c_gifProps.Width, c_gifProps.Height, _
                    c_BkBuff.DC, mirrorX, mirrorY, mirrorCx, mirrorCy, vbSrcCopy
                SelectObject c_BkBuff.DC, c_BkBuff.hDib
            End If
            
        End If
    End If
    
End Sub

Private Sub ScaleToDestination(ByRef Cx As Long, ByRef Cy As Long)

    ' function scales an image to the target destination based on the
    ' stretch mode setting, image dimensions, and target dimensions
    
    If c_curFrame = 0& Then Exit Sub ' no gif loaded yet
    
    Dim oCx As Long, oCy As Long
    oCx = UserControl.ScaleWidth
    oCy = UserControl.ScaleHeight
    
    If (c_ScaleMode And Not gfsAutoSize) = gfsClip Then
        c_gifProps.ScaleCx = 1!: c_gifProps.ScaleCy = 1!
        Cx = oCx: Cy = oCy
    
    ElseIf (c_ScaleMode And Not gfsAutoSize) = gfsActualSize Then
        c_gifProps.ScaleCx = 1!: c_gifProps.ScaleCy = 1!
        Cx = c_gifProps.Width
        Cy = c_gifProps.Height
    
    Else ' scaling in one way or another....
        c_gifProps.ScaleCx = oCx / c_gifProps.Width
        c_gifProps.ScaleCy = oCy / c_gifProps.Height
        Select Case (c_ScaleMode And Not gfsAutoSize)
        Case gfsStretch
            ' nothing else to calculate when stretching
        Case gfsShrinkScaleToFit
            If c_gifProps.ScaleCx > 1! And c_gifProps.ScaleCy > 1! Then
                c_gifProps.ScaleCx = 1! ' image will fit without scaling; use 1:1 scaling
                c_gifProps.ScaleCy = 1!
            Else            ' image must be scaled; use same ratio for width/height
                If c_gifProps.ScaleCx > c_gifProps.ScaleCy Then c_gifProps.ScaleCx = c_gifProps.ScaleCy Else c_gifProps.ScaleCy = c_gifProps.ScaleCx
            End If
        Case Else ' always scale to target dimensions; use same ratio for width/height
            If c_gifProps.ScaleCx > c_gifProps.ScaleCy Then c_gifProps.ScaleCx = c_gifProps.ScaleCy Else c_gifProps.ScaleCy = c_gifProps.ScaleCx
        End Select
    
        If (c_ScaleMode And gfsAutoSize) Then
            Cx = c_gifProps.Width * c_gifProps.ScaleCx
            Cy = c_gifProps.Height * c_gifProps.ScaleCy
        Else
            Cx = oCx
            Cy = oCy
        End If
    End If
End Sub

Private Function PictureFromByteStream(inArray() As Byte, Size As Long) As IPicture
  
  ' function creates a stdPicture (IPicture interface) from a byte array
  ' NOTE: This is a very unforgiving function. If the stream is not in the proper format
  '       the OleLoadPicture API will most likely lock up the application (GPF)
  ' -- Don't modify this routine or the routines that call this function
  
    Dim o_hMem  As Long
    Dim o_lpMem  As Long
    Dim aGUID(0 To 3) As Long
    Dim IIStream As IUnknown
    
    aGUID(0) = &H7BF80980    ' GUID for stdPicture
    aGUID(1) = &H101ABF32
    aGUID(2) = &HAA00BB8B
    aGUID(3) = &HAB0C3000
    
    o_hMem = GlobalAlloc(&H2&, Size)
    If Not o_hMem = 0& Then
        o_lpMem = GlobalLock(o_hMem)
        If Not o_lpMem = 0& Then
            CopyMemory ByVal o_lpMem, inArray(LBound(inArray)), Size
            Call GlobalUnlock(o_hMem)
            If CreateStreamOnHGlobal(o_hMem, 1&, IIStream) = 0& Then
                  Call OleLoadPicture(ByVal ObjPtr(IIStream), 0&, 0&, aGUID(0), PictureFromByteStream)
            End If
        End If
    End If
End Function

Private Function ConvertStripToGIF(ByVal Index As Long) As StdPicture

'--- Return as GIF; need to LZW compress bitmap bytes into GIF format
    
' HEAVILY MODIFIED AUTHOR'S ORIGINAL CODE (by Carles P.V.)
' - Reorganized the original routine; removed GOTOs and flow is smoother in hash routine
' - Pulled 15 global declarations into this routine.
' - The following original routines were combined into this one:
'   pvClearBlock, pvClearTable, pvCharInit & InitMasks
' - Overall changes might be a hair slower for small gifs but faster for bigger gifs
'   example using Win98's Clouds.BMP (8bit, 640x480); in IDE: 160 ms faster on my pc

  Const MAX_BITS                    As Long = 12    ' Per GIF docs, 12 is the Max
  Const MAX_BITSHIFT                As Long = 4096  ' 2^MAXBITS
  Const MAX_CODE                    As Long = 4096  ' Should NEVER generate this code
  Const TABLE_SIZE                  As Long = 5003  ' 80% occupancy (hash)
  
  Dim LIdx     As Long      ' hash table index
  Dim lFCode   As Long      ' pixel pattern (hashed)
  Dim lC       As Long      ' most recent pixel
  Dim lEnt     As Long      ' previous known pattern
  Dim lDisp    As Long
  Dim m_lShift As Long
  
  Dim m_lCodeCount As Long      ' count of codes used
  Dim m_lMaxCode As Long        ' Maximum code, given m_lBits
  Dim m_BitBucketCount As Long  ' bit bucket bit counter
  Dim m_BitBucketBuff As Long   ' cache of bits/bytes to process
  Dim m_lCurrentBits As Long    ' current lzw compression size (variable)
  Dim m_lSubBlockSize As Long   ' flag to track last byte in a data subBlock

  '-- Block compression parameters.
  Dim m_lInitBits      As Long  ' baseline LZW compression size
  Dim m_lClearCode     As Long  ' clear code
  Dim m_lHashTable() As Long
  Dim m_lCodeTable() As Long

  On Error GoTo EH
  ReDim m_lHashTable(0 To TABLE_SIZE - 1&)
  ReDim m_lCodeTable(0 To TABLE_SIZE - 1&)

  ' Added by LaVolpe to read from the custom DIB strip
  Dim tSA As SAFEARRAY2D, dibBytes() As Byte        ' DMA DIB processing
  Dim bExistingPattern As Boolean                   ' indicates hash match found
  Dim dibRow As Long, dibCol As Long, aPtr As Long  ' DIB looping parameters
  Dim tBMPI As BITMAPINFO
  
    'Added by LaVolpe -- gif frame custom build; byte at a time
    ' the actual width & height of the DIB strip is not cached; get it
    tBMPI.bmiHeader.biSize = 40
    GetDIBits c_DC.DC, c_DC.hDib, 0&, 0&, ByVal 0&, tBMPI, 0&
    With tSA
        .rgSABound(0).cElements = tBMPI.bmiHeader.biHeight
        .rgSABound(1).cElements = ByteAlignOnWord(8, tBMPI.bmiHeader.biWidth)
        .cbElements = 1
        .cDims = 2
        .pvData = c_DC.dibPtr
    End With
    CopyMemory ByVal VarPtrArray(dibBytes()), VarPtr(tSA), 4&
    
    ' oversize/guesstimate compressed GIF data & include GIF block information
    ReDim c_DIBarray(0 To 800& + (c_Frames(Index).Dimensions.Bottom * c_Frames(Index).Dimensions.Right))
    '^ 800= 13 hdr + 10 img descrip + 8 img ctrl + 768 palette + 1 EOF flag
    With c_Frames(Index)
        ' start building the GIF frame by hand, a byte at a time
        CopyMemory c_DIBarray(0), &H38464947, 4&  ' add the 6-byte GIF89a signature
        CopyMemory c_DIBarray(4), &H6139&, 2&
        CopyMemory c_DIBarray(6), .Dimensions.Right, 2&   ' now the width
        CopyMemory c_DIBarray(8), .Dimensions.Bottom, 2&  ' & height
        ' add the color table flag and the table itself
        c_DIBarray(10) = 135 ' global color table @ 8 bits (128 or 7)
        ' can skip next 2 bytes: bkg window palette index & pixel aspect ratio
        aPtr = 13&
        For LIdx = 0& To c_ColorTables.Tables(PALETTECOUNT, .TblIndex) - 1& ' convert BGR palette to RGB
            c_DIBarray(aPtr) = (c_ColorTables.Tables(LIdx, .TblIndex) \ &H10000) And &HFF
            c_DIBarray(aPtr + 1&) = (c_ColorTables.Tables(LIdx, .TblIndex) \ &H100&) And &HFF
            c_DIBarray(aPtr + 2&) = c_ColorTables.Tables(LIdx, .TblIndex) And &HFF
            aPtr = aPtr + 3
        Next
        aPtr = 781& ' 256 * 3 + 13 : note that aPtr can come out of loop less than 781
        ' add the image control block if needed
        If .IsTransparent = 1 Then
            CopyMemory c_DIBarray(aPtr), &H4F921, 3&    ' Introducer(33);BlockID(249);BlockLen(4)
            c_DIBarray(aPtr + 3&) = 1                   ' transparency flag
            ' skip next 2 bytes which is the Delay time; not needed for single frame
            c_DIBarray(aPtr + 6&) = .TransIndex
            ' skip next byte which is the block terminator - always zero
            aPtr = aPtr + 8& ' next position in array
        End If
        ' add the image description block
        c_DIBarray(aPtr) = 44 ' BlockID
        ' Left and Top coords are zeros; write frame width & height
        CopyMemory c_DIBarray(aPtr + 5&), c_DIBarray(6), 4&
        ' the packed byte would be next; but we'll leave it as zero
        ' since frame is not interlaced and not using a local color table.
        c_DIBarray(aPtr + 10&) = 8 ' LZW compression size; image bit depth
        aPtr = aPtr + 11& '11=img description block size of 10 + LZW compression byte
        ' next comes compressing DIB into LZW sub blocks, then finishing off stream
    End With
    
    'Initialize Masks -- Init LUT for fast 2 ^ x - 1 (was InitMasks routine)
    ReDim c_aPOT(0 To 16)
    c_aPOT(0) = 0
    For LIdx = 1& To 16&
        c_aPOT(LIdx) = 2 * (c_aPOT(LIdx - 1&) + 1&) - 1&
    Next LIdx
    
    '-- Reset output buffer values
    ReDim c_aBuff(0 To 254) ' (was pvCharInit)

    '-- Set up the necessary startup values
    m_lInitBits = 9& '(DIB bit depth + 1)
    m_lCurrentBits = m_lInitBits
    m_lMaxCode = c_aPOT(m_lCurrentBits)
    m_lClearCode = PALETTECOUNT     ' 2^(m_lInitBits - 1)
    m_lCodeCount = PALETTECOUNT + 2& ' 2^(m_lInitBits - 1) + 2
    
    '-- Set hash code range bound for shifting
    lFCode = TABLE_SIZE
    Do While lFCode < 65536
        m_lShift = m_lShift + 1&
        lFCode = lFCode + lFCode
    Loop
    m_lShift = 1& + c_aPOT(8& - m_lShift)
    'Added by LaVolpe -- quick erase, setting all table entries to -1 (was pvClearTable)
    FillMemory m_lHashTable(0), TABLE_SIZE * 4&, 255  ' clear hash table
    
    '-- Start...
    m_lSubBlockSize = 1& ' position for 1st byte in data sub block
    ' all images begin with a clear table flag
    Call pvOutputCode(m_lClearCode, aPtr, m_BitBucketCount, m_BitBucketBuff, m_lCurrentBits, m_lSubBlockSize)
    
    ' start LZW patterns & also start looping on 2nd pixel
    lEnt = dibBytes(c_Frames(Index).imgOffset, c_Frames(Index).Dimensions.Bottom - 1&)
    dibCol = c_Frames(Index).imgOffset + 1&
    
    'Added by LaVolpe -- looping is my enhancement;
    ' hash algorithm reorganized, comments added, modified very little
    For dibRow = tBMPI.bmiHeader.biHeight - 1& To tBMPI.bmiHeader.biHeight - c_Frames(Index).Dimensions.Bottom Step -1&
    
        ' process each pixel in line of image
        For dibCol = dibCol To c_Frames(Index).imgOffset + c_Frames(Index).Dimensions.Right - 1&
        
            lC = dibBytes(dibCol, dibRow)
            
            lFCode = lC * MAX_BITSHIFT + lEnt   ' add to existing pattern
            LIdx = (lC * m_lShift) Xor lEnt     ' XOR hashing
    
            If LIdx >= TABLE_SIZE Then LIdx = 0& ' added by LaVolpe (sanity check)
            
            If (m_lHashTable(LIdx) = lFCode) Then   ' found existing pattern
                lEnt = m_lCodeTable(LIdx)
            Else
                If (m_lHashTable(LIdx) > -1&) Then ' else Empty slot
                    
                    lDisp = TABLE_SIZE - LIdx     ' Secondary hash (after G. Knott)
                    If (LIdx = 0&) Then lDisp = 1&
                    
                    Do  ' Hash Probing
                        LIdx = LIdx - lDisp
                        If (LIdx < 0&) Then LIdx = LIdx + TABLE_SIZE
        
                        If (m_lHashTable(LIdx) = lFCode) Then
                            lEnt = m_lCodeTable(LIdx)
                            bExistingPattern = True
                            Exit Do
                        End If
        
                    Loop While (m_lHashTable(LIdx) > 0&) ' Continue probing
                End If
                
                If bExistingPattern = True Then
                    bExistingPattern = False    ' reset flag
                Else
                    ' write previous pattern
                    Call pvOutputCode(lEnt, aPtr, m_BitBucketCount, m_BitBucketBuff, m_lCurrentBits, m_lSubBlockSize)
                    lEnt = lC  ' set current palette index as previous pattern
                    m_lCodeTable(LIdx) = m_lCodeCount ' store code & hash index
                    m_lHashTable(LIdx) = lFCode
                    
                    ' check for LZW compression increments
                    If m_lCodeCount > m_lMaxCode Then
                        ' ran out of codes for current compression size
                        If (m_lCodeCount = MAX_CODE) Then
                            ' add clear code to output stream (was pvClearBlock)
                            Call pvOutputCode(m_lClearCode, aPtr, m_BitBucketCount, m_BitBucketBuff, m_lCurrentBits, m_lSubBlockSize)
                            FillMemory m_lHashTable(0), TABLE_SIZE * 4&, 255  ' clear hash table
                            m_lCurrentBits = m_lInitBits    ' reset baseline LZW compression size
                            m_lCodeCount = PALETTECOUNT + 1& ' reset code counter to 1 less 'cause it is incremented right away
                        Else
                            m_lCurrentBits = m_lCurrentBits + 1& ' increment compression size
                        End If
                        m_lMaxCode = c_aPOT(m_lCurrentBits) ' new max count for current compression size
                    End If
                    m_lCodeCount = m_lCodeCount + 1&  ' increment the number of patterns
                End If
            End If
        Next dibCol
        dibCol = c_Frames(Index).imgOffset  ' reset to 1st pixel of line
    Next dibRow

    '--  Put out the final code & image data termination code
    Call pvOutputCode(lEnt, aPtr, m_BitBucketCount, m_BitBucketBuff, m_lCurrentBits, m_lSubBlockSize)
    ' finish off the stream & add end of image data flag
    Call pvOutputCode(PALETTECOUNT + 1&, aPtr, m_BitBucketCount, m_BitBucketBuff, m_lCurrentBits, m_lSubBlockSize)
    ' flush remaining bytes in output bitbucket
    Call pvOutputCode(-1, aPtr, m_BitBucketCount, m_BitBucketBuff, m_lCurrentBits, m_lSubBlockSize)
    ' the next byte would be a sub block terminator which is zero; ignore & go on
    c_DIBarray(aPtr + 1) = 59   ' add end of file flag
        
    Set ConvertStripToGIF = PictureFromByteStream(c_DIBarray(), aPtr + 2&)

EH:
    If Err Then
        Stop        ' for testing if error occurs
        Err.Clear
        Resume      ' for testing if error occurs
    End If
    
    ' clear global arrays, no longer used
    If tSA.cbElements = 1 Then CopyMemory ByVal VarPtrArray(dibBytes()), 0&, 4&
    Erase c_DIBarray()
    Erase c_aPOT()
    Erase c_aBuff()
End Function

Private Function isStringANSI(inText As String) As Boolean

    ' simple test to determine if passed string is ANSI-like or not.
    ' In other words, does it contain unicode characters.
    
    Dim tArray() As Byte
    Dim X As Long
    
    If inText = vbNullString Then
    
        isStringANSI = True
    
    Else
    
        tArray = inText
        For X = LBound(tArray) + 1 To UBound(tArray) Step 2
            If Not tArray(X) = 0 Then Exit For
        Next
        
        isStringANSI = (X > UBound(tArray))
        
    End If

End Function


Private Sub pvOutputCode(ByVal lCode As Long, ByRef arrayPtr As Long, _
                        lBitBucketCount As Long, lBitBucketBuff As Long, _
                        lCurrentBits As Long, lSubBlockSize As Long)


' MODIFIED/REORGANIZED AUTHOR'S ORIGINAL CODE (by Carles P.V.)
' - Also combined original pvCharOut and pvFlushChar routines herein

    If lCode < 0 Then
    
        If Not lBitBucketCount = 0 Then ' add last bits of the output buffer to array
            c_aBuff(lSubBlockSize) = lBitBucketBuff
            lSubBlockSize = lSubBlockSize + 1
        End If
        If Not lSubBlockSize = 1 Then ' still have bytes to write
            c_aBuff(0) = lSubBlockSize
            ' ensure array large enough for this block + 5 extra bytes needed to finish off stream
            If UBound(c_DIBarray) < arrayPtr + lSubBlockSize + 6 Then
                ReDim Preserve c_DIBarray(0 To arrayPtr + lSubBlockSize + 6)
            End If
            CopyMemory c_DIBarray(arrayPtr), c_aBuff(0), lSubBlockSize + 1
            arrayPtr = arrayPtr + lSubBlockSize + 1
        End If
    
    Else
        ' add latest code to the bitbucket & track total bit count
        lBitBucketBuff = lBitBucketBuff Or (lCode * (c_aPOT(lBitBucketCount) + 1))
        lBitBucketCount = lBitBucketCount + lCurrentBits
    
        Do Until lBitBucketCount < 8
            ' remove 8 bits at a time and place in output buffer (byte array)
            c_aBuff(lSubBlockSize) = (lBitBucketBuff And &HFF&)
            lBitBucketBuff = lBitBucketBuff \ &H100&
            lBitBucketCount = lBitBucketCount - 8
            If (lSubBlockSize = 254) Then
                ' max allowable subblock data size is 255 byte blocks; write it
                c_aBuff(0) = lSubBlockSize
                If UBound(c_DIBarray) < arrayPtr + 260 Then
                    ' sanity check; unless the image is not compressible we
                    ' shouldn't trigger this Redim. ^^ 260=255+5 end of file/image bytes
                    ReDim Preserve c_DIBarray(0 To arrayPtr + 512)
                End If
                CopyMemory c_DIBarray(arrayPtr), c_aBuff(0), lSubBlockSize + 1
                arrayPtr = arrayPtr + lSubBlockSize + 1
                lSubBlockSize = 0
            End If
            ' keep track of next byte position to write to
            lSubBlockSize = lSubBlockSize + 1
        Loop
    End If

End Sub

Private Function BuildBackBuffer(IncludeMask As Boolean) As Boolean

    Dim dDC As Long, bErrors As Boolean, tBMPI As BITMAPINFO
    ' Simply creates a 24bpp DIB of the overall GIF dimensions
    With tBMPI.bmiHeader
        .biSize = 40
        .biBitCount = 24
        .biHeight = c_gifProps.Height
        .biWidth = c_gifProps.Width
        .biPlanes = 1
    End With
    
    ' if any errors happen, cleanup occurs in the calling routine: LoadGIF
    dDC = GetDC(0&)
    c_BkBuff.DC = CreateCompatibleDC(dDC)
    If c_BkBuff.DC = 0& Then
        bErrors = True
    Else
        c_BkBuff.hDib = CreateDIBSection(dDC, tBMPI, 0, c_BkBuff.dibPtr, 0, 0)
        If c_BkBuff.hDib = 0& Then
            bErrors = True
        Else
            c_BkBuff.hBmp = SelectObject(c_BkBuff.DC, c_BkBuff.hDib)
            If IncludeMask Then
                ' the buffer requires a mask, create that too
                With tBMPI.bmiHeader
                    .biBitCount = 1         ' monochrome mask
                    .biClrUsed = 2
                    .biClrImportant = 2
                End With
                tBMPI.bmiPalette(1) = vbWhite
                c_BkBuff.hDibBW = CreateDIBSection(dDC, tBMPI, 0, c_BkBuff.dibPtrBW, 0, 0)
                bErrors = (c_BkBuff.hDibBW = 0&)
            End If
        End If
    End If
    ReleaseDC 0&, dDC
    BuildBackBuffer = Not bErrors
End Function

Private Function IsInUserMode() As Boolean
    Dim bResult As Boolean
    On Error Resume Next
    bResult = UserControl.Ambient.UserMode
    If Err Then
        Err.Clear
    Else
        IsInUserMode = bResult
    End If
End Function

Private Function ManageTimers(ByVal StartTimer As Boolean, DelayMode As Boolean) As Boolean

    If IsInUserMode = False Then Exit Function

    If StartTimer = False Then
        If c_tmrID Then KillTimer c_tmrOwner, c_tmrID
        c_tmrID = 0&
        c_aniState = gfaStop
        ManageTimers = True
    Else
        Dim newDelay As Long
        If Not (c_aniState = gfaPlay) Then
            c_aniState = gfaStop
            If Not c_curFrame = 0& Then
                If UBound(c_Frames) > 1 Then        ' only if we have more than 1 frame
                    If c_Ptr = 0& Then
                        c_Ptr = zb_AddressOf(1, 4)
                        If c_Ptr = 0& Then
                            ReDim Preserve c_Frames(1 To 1)
                            c_Frames(1).Disposal = 0&
                            Exit Function
                        End If
                    End If
                    c_tmrOwner = UserControl.ContainerHwnd
                    If DelayMode Then       ' delay startup timer
                        c_tmrID = -ObjPtr(Me)
                        SetTimer c_tmrOwner, c_tmrID, 20, c_Ptr
                    Else                ' runtime timer
                        If c_StepDelay Then
                            newDelay = c_StepDelay
                        Else
                            newDelay = c_Frames(c_curFrame).Delay
                            If newDelay < c_MinDelay Then newDelay = c_MinDelay
                        End If
                        c_tmrID = ObjPtr(Me)
                        SetTimer c_tmrOwner, c_tmrID, newDelay, c_Ptr
                        c_aniState = gfaPlay
                    End If
                    ManageTimers = True
                End If
            End If
        End If
    End If
End Function

'*************************************************************************************************
'* cCallback - Class generic callback template
'*
'* Note:
'*  The callback declarations and code are exactly the same for a Class, Form or UserControl.
'*  The callback declarations and code can co-exist with subclassing declarations and code.
'*    With both types of code in a single file,..
'*      delete the duplicated declarations and code, Ctrl+F5 will find them for you
'*      pay careful attention to the nOrdinal parameter to zAddressOf
'*
'* Paul_Caton@hotmail.com
'* Copyright free, use and abuse as you see fit.
'*
'* v1.0 The original..................................................................... 20060408
'* v1.1 Added multi-thunk support........................................................ 20060409
'* v1.2 Added optional IDE protection.................................................... 20060411
'* v1.3 Added an optional callback target object......................................... 20060413
'*************************************************************************************************

'-Callback code-----------------------------------------------------------------------------------
Private Function zb_AddressOf(ByVal nOrdinal As Long, _
                              ByVal nParamCount As Long, _
                     Optional ByVal nThunkNo As Long = 0, _
                     Optional ByVal oCallback As Object = Nothing, _
                     Optional ByVal bIdeSafety As Boolean = True) As Long   'Return the address of the specified callback thunk
'*************************************************************************************************
'* nOrdinal     - Callback ordinal number, the final private method is ordinal 1, the second last is ordinal 2, etc...
'* nParamCount  - The number of parameters that will callback
'* nThunkNo     - Optional, allows multiple simultaneous callbacks by referencing different thunks... adjust the MAX_THUNKS Const if you need to use more than two thunks simultaneously
'* oCallback    - Optional, the object that will receive the callback. If undefined, callbacks are sent to this object's instance
'* bIdeSafety   - Optional, set to false to disable IDE protection.
'*************************************************************************************************
Const MAX_FUNKS   As Long = 1                                               'Number of simultaneous thunks, adjust to taste
Const FUNK_LONGS  As Long = 22                                              'Number of Longs in the thunk
Const FUNK_LEN    As Long = FUNK_LONGS * 4                                  'Bytes in a thunk
Const MEM_LEN     As Long = MAX_FUNKS * FUNK_LEN                            'Memory bytes required for the callback thunk
Const PAGE_RWX    As Long = &H40&                                           'Allocate executable memory
Const MEM_COMMIT  As Long = &H1000&                                         'Commit allocated memory
  Dim nAddr       As Long
  Dim z_CB() As Long
  
    If nThunkNo < 0 Or nThunkNo > (MAX_FUNKS - 1) Then
      MsgBox "nThunkNo doesn't exist.", vbCritical + vbApplicationModal, "Error in " & TypeName(Me) & ".cb_Callback"
      Exit Function
    End If
    
  If z_CbMem = 0& Then                                                  'If memory hasn't been allocated
    If oCallback Is Nothing Then                                              'If the user hasn't specified the callback owner
      Set oCallback = Me                                                      'Then it is me
    End If
  
    nAddr = zAddressOf(oCallback, nOrdinal)                                   'Get the callback address of the specified ordinal
    If nAddr = 0 Then
      MsgBox "Callback address not found.", vbCritical + vbApplicationModal, "Error in " & TypeName(Me) & ".cb_Callback"
      Exit Function
    End If
    
    z_CbMem = VirtualAlloc(z_CbMem, MEM_LEN, MEM_COMMIT, PAGE_RWX)          'Allocate executable memory
    If z_CbMem = 0& Then Exit Function
    ReDim z_CB(0 To FUNK_LONGS - 1, 0 To MAX_FUNKS - 1) As Long             'Create the machine-code array
    
    If z_CB(0, nThunkNo) = 0 Then                                             'If this ThunkNo hasn't been initialized...
      z_CB(3, nThunkNo) = _
                GetProcAddress(GetModuleHandleA("kernel32"), "IsBadCodePtr")
      z_CB(4, nThunkNo) = &HBB60E089
      z_CB(5, nThunkNo) = z_CbMem                                             'Set the data address
      z_CB(6, nThunkNo) = &H73FFC589: z_CB(7, nThunkNo) = &HC53FF04: z_CB(8, nThunkNo) = &H7B831F75: z_CB(9, nThunkNo) = &H20750008: z_CB(10, nThunkNo) = &HE883E889: z_CB(11, nThunkNo) = &HB9905004: z_CB(13, nThunkNo) = &H74FF06E3: z_CB(14, nThunkNo) = &HFAE2008D: z_CB(15, nThunkNo) = &H53FF33FF: z_CB(16, nThunkNo) = &HC2906104: z_CB(18, nThunkNo) = &H830853FF: z_CB(19, nThunkNo) = &HD87401F8: z_CB(20, nThunkNo) = &H4589C031: z_CB(21, nThunkNo) = &HEAEBFC
    End If
    
    z_CB(0, nThunkNo) = ObjPtr(oCallback)                                     'Set the Owner
    z_CB(1, nThunkNo) = nAddr                                                 'Set the callback address
    
    If bIdeSafety Then                                                        'If the user wants IDE protection
      z_CB(2, nThunkNo) = GetProcAddress(GetModuleHandleA("vba6"), "EbMode")  'EbMode Address
    End If
      
    z_CB(12, nThunkNo) = nParamCount                                          'Set the parameter count
    z_CB(17, nThunkNo) = nParamCount * 4                                      'Set the number of stck bytes to release on thunk return
    
    nAddr = z_CbMem + (nThunkNo * FUNK_LEN)                                   'Calculate where in the allocated memory to copy the thunk
    RtlMoveMemory nAddr, VarPtr(z_CB(0, nThunkNo)), FUNK_LEN                  'Copy thunk code to executable memory
    zb_AddressOf = nAddr + 16                                                 'Thunk code start address
  Else
    nAddr = z_CbMem + (nThunkNo * FUNK_LEN)
  End If
End Function

'Return the address of the specified ordinal method on the oCallback object, 1 = last private method, 2 = second last private method, etc
Private Function zAddressOf(ByVal oCallback As Object, ByVal nOrdinal As Long) As Long
  Dim bSub  As Byte                                                         'Value we expect to find pointed at by a vTable method entry
  Dim bVal  As Byte
  Dim nAddr As Long                                                         'Address of the vTable
  Dim I     As Long                                                         'Loop index
  Dim J     As Long                                                         'Loop limit
  
  RtlMoveMemory VarPtr(nAddr), ObjPtr(oCallback), 4                         'Get the address of the callback object's instance
  If Not zProbe(nAddr + &H1C, I, bSub) Then                                 'Probe for a Class method
    If Not zProbe(nAddr + &H6F8, I, bSub) Then                              'Probe for a Form method
      If Not zProbe(nAddr + &H7A4, I, bSub) Then                            'Probe for a UserControl method
        Exit Function                                                       'Bail...
      End If
    End If
  End If
  
  I = I + 4                                                                 'Bump to the next entry
  J = I + 1024                                                              'Set a reasonable limit, scan 256 vTable entries
  Do While I < J
    RtlMoveMemory VarPtr(nAddr), I, 4                                       'Get the address stored in this vTable entry
    
    If IsBadCodePtr(nAddr) Then                                             'Is the entry an invalid code address?
      RtlMoveMemory VarPtr(zAddressOf), I - (nOrdinal * 4), 4               'Return the specified vTable entry address
      Exit Do                                                               'Bad method signature, quit loop
    End If

    RtlMoveMemory VarPtr(bVal), nAddr, 1                                    'Get the byte pointed to by the vTable entry
    If bVal <> bSub Then                                                    'If the byte doesn't match the expected value...
      RtlMoveMemory VarPtr(zAddressOf), I - (nOrdinal * 4), 4               'Return the specified vTable entry address
      Exit Do                                                               'Bad method signature, quit loop
    End If
    
    I = I + 4                                                             'Next vTable entry
  Loop
End Function

'Probe at the specified start address for a method signature
Private Function zProbe(ByVal nStart As Long, ByRef nMethod As Long, ByRef bSub As Byte) As Boolean
  Dim bVal    As Byte
  Dim nAddr   As Long
  Dim nLimit  As Long
  Dim nEntry  As Long
  
  nAddr = nStart                                                            'Start address
  nLimit = nAddr + 32                                                       'Probe eight entries
  Do While nAddr < nLimit                                                   'While we've not reached our probe depth
    RtlMoveMemory VarPtr(nEntry), nAddr, 4                                  'Get the vTable entry
    
    If nEntry <> 0 Then                                                     'If not an implemented interface
      RtlMoveMemory VarPtr(bVal), nEntry, 1                                 'Get the value pointed at by the vTable entry
      If bVal = &H33 Or bVal = &HE9 Then                                    'Check for a native or pcode method signature
        nMethod = nAddr                                                     'Store the vTable entry
        bSub = bVal                                                         'Store the found method signature
        zProbe = True                                                       'Indicate success
        Exit Function                                                       'Return
      End If
    End If
    
    nAddr = nAddr + 4                                                       'Next vTable entry
  Loop
End Function

Private Sub zTerminate()
    
    Const MEM_RELEASE As Long = &H8000&                                'Release allocated memory flag
    If Not z_CbMem = 0 Then                                            'If memory allocated
        If Not VirtualFree(z_CbMem, 0, MEM_RELEASE) = 0 Then
            z_CbMem = 0  'Release; Indicate memory released
        End If
    End If
End Sub

'*************************************************************************************************
'* Callback - the final private routine is ordinal #1, second last is ordinal #2 etc
'*************************************************************************************************
'Callback ordinal 1
Private Function TimerProc(ByVal hWnd As Long, ByVal tMsg As Long, ByVal TimerID As Long, ByVal tickCount As Long) As Long
    
    KillTimer c_tmrOwner, TimerID    ' stop current timer
    
    Dim bRestart As Boolean
    Dim bLoopComplete As Boolean
    Dim tValue As Long
    
    If TimerID = ObjPtr(Me) Then
    
        ' determine next frame in the animation order
        tValue = c_curFrame
        c_curFrame = c_curFrame + 1
        If c_curFrame > UBound(c_Frames) Then
            c_curFrame = 1
            bLoopComplete = True
        End If
        If UBound(c_Frames) = 1 Then  ' single frame gif
            c_AniLoops = 0            ' shouldn't get here/timer should not have been created
            c_aniState = gfaStop
        Else
            ' determine if timer should continue
            bRestart = True
            If bLoopComplete = True Then ' another loop finished
                If Not c_gifProps.Loops = 0 Then  ' has specified number of animation loops
                    c_AniLoops = c_AniLoops - 1 ' decrease number of loops remaining
                    If c_AniLoops = 0 Then ' end of desired loops, no more animation
                        c_aniState = gfaStop
                        c_AniLoops = c_gifProps.Loops   ' reset internal counter
                        c_curFrame = UBound(c_Frames)
                        bRestart = False
                    End If
                End If
            End If
        End If
        
        UserControl.Refresh
        RaiseEvent FrameChanged(tValue, True)
    
        If bRestart = True Then ' set timer for current frame
            If c_StepDelay Then
                tValue = c_StepDelay
            Else
                If c_Frames(c_curFrame).Delay < c_MinDelay Then
                    tValue = c_MinDelay
                Else
                    tValue = c_Frames(c_curFrame).Delay
                End If
            End If
            SetTimer hWnd, TimerID, tValue, c_Ptr ' set timer
        Else
            RaiseEvent LoopsEnded
        End If

    Else ' continue processing the remaining frames

        c_tmrID = 0&
        BuildDIBstrip False
        
    End If
EH:
' CAUTION: DO NOT ADD ANY ADDITIONAL CODE OR COMMENTS PAST THE "END FUNCTION"
'          STATEMENT BELOW. Paul Caton's zProbe routine will read it as a start
'          of a new function/sub and the class timer's will fail every time.
End Function
