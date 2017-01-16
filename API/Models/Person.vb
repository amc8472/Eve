Public Class Person



    Private PersonIdValue As Integer
    Private departmentValue As Integer

        Public Sub New()
        End Sub

        Public Sub New(id As Integer, departmentId As Integer)
            EmployeeId = id
            Department = departmentId
        End Sub

        Public Property Department As Integer
            Get
                Return departmentValue
            End Get
            Set
                departmentValue = Value
            End Set
        End Property

        Public Property EmployeeId As Integer
            Get
                Return employeeIdValue
            End Get
            Set
                employeeIdValue = Value
            End Set
        End Property

    End Function
    End Class

