VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsTheme"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False

Option Explicit

Private Const DEFAULT_FONT_NAME As String = "Segoe UI"
Private Const DEFAULT_FONT_SIZE As Single = 10
Private Const DEFAULT_TITLE_FONT_SIZE As Single = 18
Private Const DEFAULT_SECTION_FONT_SIZE As Single = 11
Private Const DEFAULT_BUTTON_FONT_SIZE As Single = 10

Private mPrimaryColor As Long
Private mPrimaryDarkColor As Long
Private mSecondaryColor As Long
Private mBackgroundColor As Long
Private mSurfaceColor As Long
Private mTextColor As Long
Private mMutedTextColor As Long
Private mBorderColor As Long
Private mSuccessColor As Long
Private mWarningColor As Long
Private mErrorColor As Long

Private mFontName As String
Private mFontSize As Single
Private mTitleFontSize As Single
Private mSectionFontSize As Single
Private mButtonFontSize As Single

Private Sub Class_Initialize()

    mPrimaryColor = RGB(31, 78, 121)
    mPrimaryDarkColor = RGB(22, 55, 85)
    mSecondaryColor = RGB(91, 155, 213)

    mBackgroundColor = RGB(242, 244, 247)
    mSurfaceColor = RGB(255, 255, 255)

    mTextColor = RGB(35, 39, 42)
    mMutedTextColor = RGB(105, 112, 119)
    mBorderColor = RGB(210, 214, 220)

    mSuccessColor = RGB(46, 125, 50)
    mWarningColor = RGB(245, 166, 35)
    mErrorColor = RGB(198, 40, 40)

    mFontName = DEFAULT_FONT_NAME
    mFontSize = DEFAULT_FONT_SIZE
    mTitleFontSize = DEFAULT_TITLE_FONT_SIZE
    mSectionFontSize = DEFAULT_SECTION_FONT_SIZE
    mButtonFontSize = DEFAULT_BUTTON_FONT_SIZE

End Sub

Public Property Get PrimaryColor() As Long
    PrimaryColor = mPrimaryColor
End Property

Public Property Get PrimaryDarkColor() As Long
    PrimaryDarkColor = mPrimaryDarkColor
End Property

Public Property Get SecondaryColor() As Long
    SecondaryColor = mSecondaryColor
End Property

Public Property Get BackgroundColor() As Long
    BackgroundColor = mBackgroundColor
End Property

Public Property Get SurfaceColor() As Long
    SurfaceColor = mSurfaceColor
End Property

Public Property Get TextColor() As Long
    TextColor = mTextColor
End Property

Public Property Get MutedTextColor() As Long
    MutedTextColor = mMutedTextColor
End Property

Public Property Get BorderColor() As Long
    BorderColor = mBorderColor
End Property

Public Property Get SuccessColor() As Long
    SuccessColor = mSuccessColor
End Property

Public Property Get WarningColor() As Long
    WarningColor = mWarningColor
End Property

Public Property Get ErrorColor() As Long
    ErrorColor = mErrorColor
End Property

Public Property Get FontName() As String
    FontName = mFontName
End Property

Public Property Get FontSize() As Single
    FontSize = mFontSize
End Property

Public Property Get TitleFontSize() As Single
    TitleFontSize = mTitleFontSize
End Property

Public Property Get SectionFontSize() As Single
    SectionFontSize = mSectionFontSize
End Property

Public Property Get ButtonFontSize() As Single
    ButtonFontSize = mButtonFontSize
End Property
