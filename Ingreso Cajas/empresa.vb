Public Class empresas
    Private Shared empresan As String
    Private Shared nomempre As String
    Private Shared conecstr As String
    Private Shared strconec As String
    Private Shared strconol As String
    Private Shared usuario_sistema As String
    Private Shared clave_sistema As String
    Private Shared seccion_c As String

    Public Sub New()
        'empresan = "Jesus Acosta"
    End Sub

    Public Property numero() As String
        Get
            Return empresan
        End Get


        Set(ByVal Value As String)
            empresan = Value
        End Set


    End Property

    Public Property nombre() As String
        Get
            Return nomempre
        End Get


        Set(ByVal Value As String)
            nomempre = Value
        End Set
    End Property

    Public Property conexion() As String
        Get
            Return conecstr
        End Get


        Set(ByVal Value As String)
            conecstr = Value
        End Set
    End Property


    Public Property constr() As String
        Get
            Return strconec
        End Get


        Set(ByVal Value As String)
            strconec = Value
        End Set
    End Property

    Public Property conole() As String
        Get
            Return strconol
        End Get


        Set(ByVal Value As String)
            strconol = Value
        End Set
    End Property

    Public Property usuario() As String
        Get
            Return usuario_sistema
        End Get

        Set(ByVal Value As String)
            usuario_sistema = Value
        End Set
    End Property

    Public Property clave() As String
        Get
            Return clave_sistema
        End Get

        Set(ByVal Value As String)
            clave_sistema = Value
        End Set
    End Property

    Public Property seccion() As String
        Get
            Return seccion_c
        End Get

        Set(ByVal Value As String)
            seccion_c = Value
        End Set
    End Property
End Class



