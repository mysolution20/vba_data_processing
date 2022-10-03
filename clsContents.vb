Option Explicit

Private name_ As String
Private contentLoad_ As String
Private typeName_ As String
'

Public Property Get getName() As String

    getName = name_

End Property

Public Property Let letName(ByVal name As String)

    name_ = name
    
End Property

Public Property Get getContent() As String

    getContent = contentLoad_

End Property

Public Property Let letContent(ByVal contentLoad As String)

    contentLoad_ = contentLoad
    
End Property

Public Property Get getType() As String

        getType = typeName_

End Property

Public Property Let letType(ByVal typeName As String)

    typeName_ = typeName

End Property