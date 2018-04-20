VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "JsonConverterParams"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit


Private m_flip As Boolean
Private m_ignoreLineman As Boolean
Private m_ignoreDefenders As Boolean
Private m_ignoreText As Boolean
Private m_ignoreSolidLines As Boolean
Private m_ignoreDashedLines As Boolean
Private m_ignorePlayers As Boolean

Public Property Get Flip() As Boolean
    Flip = m_flip
End Property
Public Property Let Flip(Flip As Boolean)
    m_flip = Flip
End Property

Public Property Get IgnoreLineman() As Boolean
    IgnoreLineman = m_ignoreLineman
End Property
Public Property Let IgnoreLineman(IgnoreLineman As Boolean)
    m_ignoreLineman = IgnoreLineman
End Property

Public Property Get IgnoreDefenders() As Boolean
    IgnoreDefenders = m_ignoreDefenders
End Property
Public Property Let IgnoreDefenders(IgnoreDefenders As Boolean)
    m_ignoreDefenders = IgnoreDefenders
End Property

Public Property Get IgnoreText() As Boolean
    IgnoreText = m_ignoreText
End Property
Public Property Let IgnoreText(IgnoreText As Boolean)
    IgnoreText = IgnoreDefenders
End Property

Public Property Get IgnoreSolidLines() As Boolean
    IgnoreSolidLines = m_ignoreSolidLines
End Property
Public Property Let IgnoreSolidLines(ByVal IgnoreSolidLines As Boolean)
    m_ignoreSolidLines = IgnoreSolidLines
End Property

Public Property Get IgnoreDashedLines() As Boolean
    IgnoreDashedLines = m_ignoreDashedLines
End Property
Public Property Let IgnoreDashedLines(ByVal IgnoreDashedLines As Boolean)
    m_ignoreDashedLines = IgnoreDashedLines
End Property

Public Property Get IgnorePlayers() As Boolean
    IgnorePlayers = m_ignorePlayers
End Property
Public Property Let IgnorePlayers(ByVal IgnorePlayers As Boolean)
    m_ignorePlayers = IgnorePlayers
End Property


Private Sub Class_Initalize()
    m_flip = False
    m_flip = False
    m_ignoreLineman = False
    m_igoreDefenders = False
    m_ignoreText = False
    m_ignoreSolidLines = False
    m_ignoreDashedLines = False
    m_ignorePlayers = False
End Sub
    
