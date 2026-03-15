Private WithEvents m_app As Excel.Application
Private m_eventFired As Boolean
Private m_eventSheet As String
Private m_eventAddress As String

Public Sub Setup(app As Excel.Application)
    Set m_app = app
    m_eventFired = False
End Sub

Public Property Get EventFired() As Boolean
    EventFired = m_eventFired
End Property

Public Property Get EventSheet() As String
    EventSheet = m_eventSheet
End Property

Public Property Get EventAddress() As String
    EventAddress = m_eventAddress
End Property

Public Sub ResetEvent()
    m_eventFired = False
    m_eventSheet = ""
    m_eventAddress = ""
End Sub

Private Sub m_app_SheetChange(ByVal Sh As Object, ByVal Target As Range)
    m_eventFired = True
    m_eventSheet = Sh.Name
    m_eventAddress = Target.Address
End Sub
