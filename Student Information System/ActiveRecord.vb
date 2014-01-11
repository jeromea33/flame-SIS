Imports System.Data.OleDb
Imports System.IO
Imports System.Runtime.Serialization

Module ActiveRecord

    Dim connnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=|DataDirectory|\StudentInformationSystem.accdb"

    ''' <summary>
    ''' For the type safeness of the ActiveRecordClass.save(SaveModeParameterHere) method
    ''' </summary>
    ''' <remarks>Implemented</remarks>
    Enum SaveMode
        Insert
        Update
    End Enum

    <Serializable()>
    Class Classes
        Implements IComparable(Of Classes)
        Public Property ClassName As String
        Public Property ClassAdvisor As Integer
        Public Property ClassLevel As Integer
        Public Property NeedGraduation As Boolean = False

        Public Sub New(ByVal AClassName As String)
            Try
                find(AClassName)
            Catch
                ClassName = AClassName
            End Try
        End Sub

        Public Sub New()
        End Sub

        Public Sub save(ByVal Mode As SaveMode)
            If Mode = SaveMode.Insert Then
                insert()
            ElseIf Mode = SaveMode.Update Then
                update()
            End If
        End Sub

        Public Sub update() 'existing rows in the database, The same rows having the same ClassName,
            Dim connection As New OleDbConnection(connnectionString)
            Dim command As New OleDbCommand("update Classes set ClassAdvisor = @classAdvisor, ClassLevel = @classLevel, NeedGraduation = @needGraduation where ClassName = @className", connection)
            command.Parameters.Add("@classAdvisor", OleDbType.Integer).Value = ClassAdvisor
            command.Parameters.Add("@classLevel", OleDbType.Integer).Value = ClassLevel
            command.Parameters.Add("@needGraduation", OleDbType.Boolean).Value = NeedGraduation
            command.Parameters.Add("@className", OleDbType.WChar).Value = ClassName
            connection.Open()
            command.ExecuteNonQuery()
            connection.Close()
        End Sub

        Public Sub insert() 'rows that does not exist in the database YET..., Insert's job is to insert it in the table having this unique ClassName
            Dim connection As New OleDbConnection(connnectionString)
            Dim command As New OleDbCommand("insert into Classes(@className,@classAdvisor,@classLevel,@needGraduation)", connection)
            command.Parameters.Add("@className", OleDbType.WChar).Value = ClassName
            command.Parameters.Add("@classAdvisor", OleDbType.Integer).Value = ClassAdvisor
            command.Parameters.Add("@classLevel", OleDbType.Integer).Value = ClassLevel
            command.Parameters.Add("@needGraduation", OleDbType.Boolean).Value = NeedGraduation
            connection.Open()
            command.ExecuteNonQuery()
            connection.Close()
        End Sub

        Public Overrides Function ToString() As String
            Return ClassName
        End Function

        Public Shared Function find(ByVal ClassNameQuery As String) As Classes
            Dim returnval As New Classes
            Dim connection As New OleDbConnection(connnectionString)
            Dim command As New OleDbCommand("select * from Classes where ClassName = @classname", connection)
            Dim reader As OleDbDataReader
            command.Parameters.Add("@classname", OleDbType.WChar).Value = ClassNameQuery
            connection.Open()
            reader = command.ExecuteReader
            While reader.Read
                returnval.ClassName = reader.GetString(0)
                returnval.ClassAdvisor = reader.GetInt32(1)
                returnval.ClassLevel = reader.GetInt32(2)
                returnval.NeedGraduation = reader.GetBoolean(3)
            End While
            reader.Close()
            connection.Close()
            If String.IsNullOrWhiteSpace(returnval.ClassName) Then
                Throw New Exception("Class does not exist, please refer to an existing classes in the database.")
            End If
            Return returnval
        End Function

        Public Shared Function findall() As List(Of Classes)
            Dim returnVal As New List(Of Classes)
            Dim aclass As Classes
            Dim connection As New OleDbConnection(connnectionString)
            Dim command As New OleDbCommand("select * from Classes", connection)
            Dim reader As OleDbDataReader
            connection.Open()
            reader = command.ExecuteReader
            While reader.Read
                aclass = New Classes
                aclass.ClassName = reader.GetString(0)
                aclass.ClassAdvisor = reader.GetInt32(1)
                aclass.ClassLevel = reader.GetInt32(2)
                aclass.NeedGraduation = reader.GetBoolean(3)
                returnVal.Add(aclass)
            End While
            reader.Close()
            connection.Close()
            Return returnVal
        End Function

        Public Shared Function delete(ByVal ClassNameQuery As String) As Integer
            Dim x As Integer
            Dim connection As New OleDbConnection(connnectionString)
            Dim command As New OleDbCommand("delete from Classes where ClassName = @className", connection)
            command.Parameters.Add("@className", OleDbType.WChar).Value = ClassNameQuery
            connection.Open()
            x = command.ExecuteNonQuery()
            connection.Close()
            Return x
        End Function

        Public Function CompareTo1(ByVal other As Classes) As Integer Implements System.IComparable(Of Classes).CompareTo
            If ClassLevel > other.ClassLevel Then
                Return 1
            ElseIf ClassLevel < other.ClassLevel Then
                Return -1
            ElseIf ClassLevel = other.ClassLevel Then

            End If
        End Function
    End Class

    <Serializable()>
    Class Student
        Implements IComparable
        Public Property StudentNumber As Integer
        Public Property LastName As String
        Public Property FirstName As String
        Public Property MiddleName As String
        Public Property Age As Integer
        Public Property Sex As String
        Public Property Birthday As Date
        Public Property Address As String
        Public Property MotherName As String
        Public Property MotherOccupation As String
        Public Property MotherSocialStatus As String
        Public Property FatherName As String
        Public Property FatherOccupation As String
        Public Property FatherSocialStatus As String
        Public Property Religion As String
        Public Property Nationality As String
        Public Property Guardian As String
        Public Property GuardianContactNumber As String
        Public Property ParentsContactNumber As String
        Public Property ContactNumber As String
        Public Property ClassEnrolled As String
        Public Property EnrollmentStatus As String

        Public Sub New(ByVal StudentNumberQuery As Integer)
            Try
                find(StudentNumberQuery)
            Catch
                StudentNumber = StudentNumberQuery
            End Try
        End Sub

        Public Sub New()
            EnrollmentStatus = "Normal"
        End Sub

        Public Sub save(ByVal mode As SaveMode)
            If mode = SaveMode.Insert Then
                insert()
            ElseIf mode = SaveMode.Update Then
                update()
            End If
        End Sub

        Public Sub insert()
            Dim connection As New OleDbConnection(connnectionString)
            Dim command As New OleDbCommand("INSERT INTO Students VALUES(@StudentNumber,@LastName,@FirstName, @MiddleName,@Age,@Sex,@Birthday,@Address,@MotherName,@MotherOccupation,@MotherSocialStatus,@FatherName,@FatherOccupation,@FatherSocialStatus,@Religion,@Nationality,@Guardian,@GuardianContactNumber,@ParentsContactNumber,@ContactNumber,@ClassEnrolled,@EnrollmentStatus)", connection)
            command.Parameters.Add("@StudentNumber", OleDbType.Integer).Value = StudentNumber
            command.Parameters.Add("@LastName", OleDbType.WChar).Value = LastName
            command.Parameters.Add("@FirstName", OleDbType.WChar).Value = FirstName
            command.Parameters.Add("@MiddleName", OleDbType.WChar).Value = MiddleName
            command.Parameters.Add("@Age", OleDbType.Integer).Value = Age
            command.Parameters.Add("@Sex", OleDbType.WChar).Value = Sex
            command.Parameters.Add("@Birthday", OleDbType.Date).Value = Birthday
            command.Parameters.Add("@Address", OleDbType.WChar).Value = Address
            command.Parameters.Add("@MotherName", OleDbType.WChar).Value = MotherName
            command.Parameters.Add("@MotherOccupation", OleDbType.WChar).Value = MotherOccupation
            command.Parameters.Add("@MotherSocialStatus", OleDbType.WChar).Value = MotherSocialStatus
            command.Parameters.Add("@FatherName", OleDbType.WChar).Value = FatherName
            command.Parameters.Add("@FatherOccupation", OleDbType.WChar).Value = FatherOccupation
            command.Parameters.Add("@FatherSocialStatus", OleDbType.WChar).Value = FatherSocialStatus
            command.Parameters.Add("@Religion", OleDbType.WChar).Value = Religion
            command.Parameters.Add("@Nationality", OleDbType.WChar).Value = Nationality
            command.Parameters.Add("@Guardian", OleDbType.WChar).Value = Address
            command.Parameters.Add("@GuardianContactNumber", OleDbType.WChar).Value = GuardianContactNumber
            command.Parameters.Add("@ParentsContactNumber", OleDbType.WChar).Value = ParentsContactNumber
            command.Parameters.Add("@ContactNumber", OleDbType.WChar).Value = ContactNumber
            command.Parameters.Add("@ClassEnrolled", OleDbType.WChar).Value = ClassEnrolled
            command.Parameters.Add("@EnrollmentStatus", OleDbType.WChar).Value = EnrollmentStatus
            connection.Open()
            command.ExecuteNonQuery()
            connection.Close()
        End Sub

        Public Sub update()
            Dim connection As New OleDbConnection(connnectionString)
            Dim command As New OleDbCommand("Update Students SET LastName = @LastName, FirstName = @FirstName, MiddleName = @MiddleName, Age = @Age, Sex = @Sex, Birthday = @Birthday, Address = @Address, MotherName = @MotherName, MotherOccupation = @MotherOccupation, MotherSocialStatus = @MotherSocialStatus, FatherName = @FatherName, FatherOccupation = @FatherOccupation, FatherSocialStatus = @FatherSocialStatus, Religion = @Religion, Nationality = @Nationality, Guardian = @Guardian, GuardianContactNumber = @GuardianContactNumber, ParentsContactNumber = @ParentsContactNumber, ContactNumber = @ContactNumber, ClassEnrolled = @ClassEnrolled, EnrollmentStatus = @EnrollmentStatus WHERE StudentNumber = @StudentNumber", connnectionString)
            command.Parameters.Add("@LastName", OleDbType.WChar).Value = LastName
            command.Parameters.Add("@FirstName", OleDbType.WChar).Value = FirstName
            command.Parameters.Add("@MiddleName", OleDbType.WChar).Value = MiddleName
            command.Parameters.Add("@Age", OleDbType.Integer).Value = Age
            command.Parameters.Add("@Sex", OleDbType.WChar).Value = Sex
            command.Parameters.Add("@Birthday", OleDbType.Date).Value = Birthday
            command.Parameters.Add("@Address", OleDbType.WChar).Value = Address
            command.Parameters.Add("@MotherName", OleDbType.WChar).Value = MotherName
            command.Parameters.Add("@MotherOccupation", OleDbType.WChar).Value = MotherOccupation
            command.Parameters.Add("@MotherSocialStatus", OleDbType.WChar).Value = MotherSocialStatus
            command.Parameters.Add("@FatherName", OleDbType.WChar).Value = FatherName
            command.Parameters.Add("@FatherOccupation", OleDbType.WChar).Value = FatherOccupation
            command.Parameters.Add("@FatherSocialStatus", OleDbType.WChar).Value = FatherSocialStatus
            command.Parameters.Add("@Religion", OleDbType.WChar).Value = Religion
            command.Parameters.Add("@Nationality", OleDbType.WChar).Value = Nationality
            command.Parameters.Add("@Guardian", OleDbType.WChar).Value = Address
            command.Parameters.Add("@GuardianContactNumber", OleDbType.WChar).Value = GuardianContactNumber
            command.Parameters.Add("@ParentsContactNumber", OleDbType.WChar).Value = ParentsContactNumber
            command.Parameters.Add("@ContactNumber", OleDbType.WChar).Value = ContactNumber
            command.Parameters.Add("@ClassEnrolled", OleDbType.WChar).Value = ClassEnrolled
            command.Parameters.Add("@StudentNumber", OleDbType.Integer).Value = StudentNumber
            command.Parameters.Add("@EnrollmentStatus", OleDbType.WChar).Value = EnrollmentStatus
            connection.Open()
            command.ExecuteNonQuery()
            connection.Close()
        End Sub

        Public Overrides Function ToString() As String
            Return StudentNumber & ": " & LastName & " " & FirstName & " " & MiddleName
        End Function

        Public Function getClassEnrolled() As Classes
            Return Classes.find(ClassEnrolled)
        End Function

        Public Shared Function find(ByVal StudentNumberQuery As Integer) As Student
            Dim numberOfRowsAffected As Integer
            Dim returnval As New Student
            Dim connection As New OleDbConnection(connnectionString)
            Dim command As New OleDbCommand("select * from Students where StudentNumber = @studentNumber", connection)
            Dim reader As OleDbDataReader
            command.Parameters.Add("@studentNumber", OleDbType.WChar).Value = StudentNumberQuery
            connection.Open()
            numberOfRowsAffected = command.ExecuteNonQuery
            reader = command.ExecuteReader
            While reader.Read
                returnval = New Student
                returnval.StudentNumber = reader.GetInt32(0)
                returnval.LastName = reader.GetString(1)
                returnval.FirstName = reader.GetString(2)
                returnval.MiddleName = reader.GetString(3)
                returnval.Age = reader.GetInt32(4)
                returnval.Sex = reader.GetString(5)
                returnval.Birthday = reader.GetDateTime(6)
                returnval.Address = reader.GetString(7)
                returnval.MotherName = reader.GetString(8)
                returnval.MotherOccupation = reader.GetString(9)
                returnval.MotherSocialStatus = reader.GetString(10)
                returnval.FatherName = reader.GetString(11)
                returnval.FatherOccupation = reader.GetString(12)
                returnval.FatherSocialStatus = reader.GetString(13)
                returnval.Religion = reader.GetString(14)
                returnval.Nationality = reader.GetString(15)
                returnval.Guardian = reader.GetString(16)
                returnval.GuardianContactNumber = reader.GetString(17)
                returnval.ParentsContactNumber = reader.GetString(18)
                returnval.ContactNumber = reader.GetString(19)
                returnval.ClassEnrolled = reader.GetString(20)
                returnval.EnrollmentStatus = reader.GetString(21)
            End While
            reader.Close()
            connection.Close()
            If numberOfRowsAffected = 0 Then
                Throw New Exception("Student does not exist. Please refer to an existing student in the database.")
            End If
            Return returnval
        End Function

        Public Shared Function findAll() As List(Of Student)
            Dim student As Student
            Dim returnval As New List(Of Student)
            Dim connection As New OleDbConnection(connnectionString)
            Dim command As New OleDbCommand("select * from Students", connection)
            Dim reader As OleDbDataReader
            connection.Open()
            reader = command.ExecuteReader
            While reader.Read
                student = New Student
                student.StudentNumber = reader.GetInt32(0)
                student.LastName = reader.GetString(1)
                student.FirstName = reader.GetString(2)
                student.MiddleName = reader.GetString(3)
                student.Age = reader.GetInt32(4)
                student.Sex = reader.GetString(5)
                student.Birthday = reader.GetDateTime(6)
                student.Address = reader.GetString(7)
                student.MotherName = reader.GetString(8)
                student.MotherOccupation = reader.GetString(9)
                student.MotherSocialStatus = reader.GetString(10)
                student.FatherName = reader.GetString(11)
                student.FatherOccupation = reader.GetString(12)
                student.FatherSocialStatus = reader.GetString(13)
                student.Religion = reader.GetString(14)
                student.Nationality = reader.GetString(15)
                student.Guardian = reader.GetString(16)
                student.GuardianContactNumber = reader.GetString(17)
                student.ParentsContactNumber = reader.GetString(18)
                student.ContactNumber = reader.GetString(19)
                student.ClassEnrolled = reader.GetString(20)
                student.EnrollmentStatus = reader.GetString(21)
                returnval.Add(student)
            End While
            connection.Close()
            Return returnval
        End Function

        Public Shared Function delete(ByVal StudentNumberQuery As Integer) As Integer
            Dim x As Integer
            Dim connection As New OleDbConnection(connnectionString)
            Dim command As New OleDbCommand("delete from Students where StudentNumber = @studentNumber", connection)
            command.Parameters.Add("@studentNumber", OleDbType.Integer).Value = StudentNumberQuery
            connection.Open()
            x = command.ExecuteNonQuery
            connection.Close()
            Return x
        End Function

        Public Function CompareTo(ByVal obj As Object) As Integer Implements System.IComparable.CompareTo
            If StudentNumber > obj.StudentNumber Then
                Return 1
            ElseIf StudentNumber < obj.StudentNumber Then
                Return -1
            Else
                Return 0
            End If
        End Function
    End Class

    <Serializable()>
    Class GuidanceRecord
        Implements IComparable
        Public Property Identifier As String
        Public Property StudentInvolved As String
        Public Property Details As String

        Public Sub New(ByVal IdentifierQuery As String)
            Try
                find(IdentifierQuery)
            Catch
                Identifier = IdentifierQuery
            End Try
        End Sub

        Public Sub New()
        End Sub

        Public Sub save(ByVal Mode As SaveMode)
            If Mode = SaveMode.Insert Then
                insert()
            ElseIf Mode = SaveMode.Update Then
                update()
            End If
        End Sub

        Public Sub insert()
            Dim connection As New OleDbConnection(connnectionString)
            Dim command As New OleDbCommand("INSERT INTO GuidanceRecords VALUES(@Identifier, @StudentInvolved, @Details)", connection)
            command.Parameters.Add("@Identifier", OleDbType.WChar).Value = Identifier
            command.Parameters.Add("@StudentInvolved", OleDbType.Integer).Value = StudentInvolved
            command.Parameters.Add("@Details", OleDbType.LongVarWChar).Value = Details
            connection.Open()
            command.ExecuteNonQuery()
            connection.Close()
        End Sub

        Public Sub update()
            Dim connection As New OleDbConnection(connnectionString)
            Dim command As New OleDbCommand("UPDATE GuidanceRecords SET StudentsInvolved = @StudentsInvolved, Details = @Details where Identifier = @Identifier", connection)
            command.Parameters.Add("@StudentInvolved", OleDbType.Integer).Value = StudentInvolved
            command.Parameters.Add("@Details", OleDbType.LongVarWChar).Value = Details
            command.Parameters.Add("@Identifier", OleDbType.WChar).Value = Identifier
            connection.Open()
            command.ExecuteNonQuery()
            connection.Close()
        End Sub

        Public Shared Function find(ByVal IdentifierQuery As String) As GuidanceRecord
            Dim returnVal As New GuidanceRecord
            Dim connection As New OleDbConnection(connnectionString)
            Dim command As New OleDbCommand("select * from GuidanceRecords where Identifier = @identifier", connection)
            Dim reader As OleDbDataReader
            command.Parameters.Add("@identifier", OleDbType.WChar).Value = IdentifierQuery
            connection.Open()
            reader = command.ExecuteReader
            While reader.Read
                returnVal.Identifier = reader.GetString(0)
                returnVal.StudentInvolved = reader.GetInt32(1)
                returnVal.Details = reader.GetString(2)
            End While
            reader.Close()
            connection.Close()
            If String.IsNullOrWhiteSpace(returnVal.Identifier) Then
                Throw New Exception("Unknown Guidance Record")
            End If
            Return returnVal
        End Function

        Public Shared Function findAll() As List(Of GuidanceRecord)
            Dim record As GuidanceRecord
            Dim returnVal As New List(Of GuidanceRecord)
            Dim connection As New OleDbConnection(connnectionString)
            Dim command As New OleDbCommand("select * from GuidanceRecords", connection)
            Dim reader As OleDbDataReader
            connection.Open()
            reader = command.ExecuteReader
            While reader.Read
                record = New GuidanceRecord
                record.Identifier = reader.GetString(0)
                record.StudentInvolved = reader.GetInt32(1)
                record.Details = reader.GetString(2)
                returnVal.Add(record)
            End While
            reader.Close()
            connection.Close()
            Return returnVal
        End Function

        Public Shared Function delete(ByVal IdentifierQuery As String) As Integer
            Dim x As Integer
            Dim connection As New OleDbConnection(connnectionString)
            Dim command As New OleDbCommand("delete from GuidanceRecords where Identifier = @Identifier", connection)
            command.Parameters.Add("@Identifier", OleDbType.WChar).Value = IdentifierQuery
            connection.Open()
            x = command.ExecuteNonQuery
            connection.Close()
            Return x
        End Function

        Public Function CompareTo(ByVal obj As Object) As Integer Implements System.IComparable.CompareTo
            If obj.Identifier = Identifier Then
                Return 0
            Else
                Return 1
            End If
        End Function
    End Class

    <Serializable()>
    Class PaymentTable
        Public Property Identifier As String
        Public Property June As Integer
        Public Property July As Integer
        Public Property August As Integer
        Public Property September As Integer
        Public Property October As Integer
        Public Property November As Integer
        Public Property December As Integer
        Public Property January As Integer
        Public Property Febuary As Integer
        Public Property March As Integer
        Public Property Graduation As Integer
        Public Property SchoolYear As Integer
        Public Property Student As Integer

        Public Sub New()
        End Sub

        Public Sub New(ByVal IdentifierQuery As String)
        End Sub

        Public Sub save(ByVal Mode As SaveMode)
            If Mode = SaveMode.Insert Then
                insert()
            ElseIf Mode = SaveMode.Update Then
                update()
            End If
        End Sub

        Public Sub insert()
            Dim connection As New OleDbConnection(connnectionString)
            Dim command As New OleDbCommand("INSERT INTO PaymentTable VALUES(@identifier,@june,@july,@august,@september,@october,@november,@december,@january,@febuary,@march,@graduation,@schoolyear,@student)", connection)
            command.Parameters.Add("@identifier", OleDbType.WChar).Value = Identifier
            command.Parameters.Add("@june", OleDbType.Integer).Value = June
            command.Parameters.Add("@july", OleDbType.Integer).Value = July
            command.Parameters.Add("@august", OleDbType.Integer).Value = August
            command.Parameters.Add("@september", OleDbType.Integer).Value = September
            command.Parameters.Add("@october", OleDbType.Integer).Value = October
            command.Parameters.Add("@november", OleDbType.Integer).Value = November
            command.Parameters.Add("@december", OleDbType.Integer).Value = December
            command.Parameters.Add("@january", OleDbType.Integer).Value = January
            command.Parameters.Add("@febuary", OleDbType.Integer).Value = Febuary
            command.Parameters.Add("@march", OleDbType.Integer).Value = March
            command.Parameters.Add("@graduation", OleDbType.Integer).Value = Graduation
            command.Parameters.Add("@schoolyear", OleDbType.Integer).Value = SchoolYear
            command.Parameters.Add("@student", OleDbType.Integer).Value = Student
            connection.Open()
            command.ExecuteNonQuery()
            connection.Close()
        End Sub

        Public Sub update()
            Dim connection As New OleDbConnection(connnectionString)
            Dim command As New OleDbCommand("UPDATE PaymentTable SET June = @june, July = @july, August = @august, September = @september, October = @october, November = @november, December = @december, January = @january, Febuary = @febuary, March = @march, Graduation = @graduation, SchoolYear = @schoolyear, Student = @student where Identifier = @identifier", connection)
            command.Parameters.Add("@june", OleDbType.Integer).Value = June
            command.Parameters.Add("@july", OleDbType.Integer).Value = July
            command.Parameters.Add("@august", OleDbType.Integer).Value = August
            command.Parameters.Add("@september", OleDbType.Integer).Value = September
            command.Parameters.Add("@october", OleDbType.Integer).Value = October
            command.Parameters.Add("@november", OleDbType.Integer).Value = November
            command.Parameters.Add("@december", OleDbType.Integer).Value = December
            command.Parameters.Add("@january", OleDbType.Integer).Value = January
            command.Parameters.Add("@febuary", OleDbType.Integer).Value = Febuary
            command.Parameters.Add("@march", OleDbType.Integer).Value = March
            command.Parameters.Add("@graduation", OleDbType.Integer).Value = Graduation
            command.Parameters.Add("@schoolyear", OleDbType.Integer).Value = SchoolYear
            command.Parameters.Add("@student", OleDbType.Integer).Value = Student
            command.Parameters.Add("@identifier", OleDbType.WChar).Value = Identifier
            connection.Open()
            command.ExecuteNonQuery()
            connection.Close()
        End Sub

        Public Shared Function find(ByVal IdentifierQuery As String) As PaymentTable
            Dim returnval As New PaymentTable
            Dim connection As New OleDbConnection(connnectionString)
            Dim command As New OleDbCommand("select * from PaymentTable where Identifier = @id", connection)
            Dim reader As OleDbDataReader
            command.Parameters.Add("@id", OleDbType.WChar).Value = IdentifierQuery
            connection.Open()
            reader = command.ExecuteReader
            While reader.Read
                returnval.Identifier = reader.GetString(0)
                returnval.June = reader.GetInt32(1)
                returnval.July = reader.GetInt32(2)
                returnval.August = reader.GetInt32(3)
                returnval.September = reader.GetInt32(4)
                returnval.October = reader.GetInt32(5)
                returnval.November = reader.GetInt32(6)
                returnval.December = reader.GetInt32(7)
                returnval.January = reader.GetInt32(8)
                returnval.Febuary = reader.GetInt32(9)
                returnval.March = reader.GetInt32(10)
                returnval.Graduation = reader.GetInt32(11)
                returnval.SchoolYear = reader.GetInt32(12)
                returnval.Student = reader.GetInt32(13)
            End While
            connection.Close()
            If String.IsNullOrWhiteSpace(returnval.Identifier) Then
                Throw New Exception("Unknown Payment Identifier")
            End If
            Return returnval
        End Function

        Public Shared Function findAll() As List(Of PaymentTable)
            Dim returnval As New List(Of PaymentTable)
            Dim val As PaymentTable
            Dim connection As New OleDbConnection(connnectionString)
            Dim command As New OleDbCommand("select * from PaymentTable where Identifier = @id", connection)
            Dim reader As OleDbDataReader
            connection.Open()
            reader = command.ExecuteReader
            While reader.Read
                val = New PaymentTable
                val.Identifier = reader.GetString(0)
                val.June = reader.GetInt32(1)
                val.July = reader.GetInt32(2)
                val.August = reader.GetInt32(3)
                val.September = reader.GetInt32(4)
                val.October = reader.GetInt32(5)
                val.November = reader.GetInt32(6)
                val.December = reader.GetInt32(7)
                val.January = reader.GetInt32(8)
                val.Febuary = reader.GetInt32(9)
                val.March = reader.GetInt32(10)
                val.Graduation = reader.GetInt32(11)
                val.SchoolYear = reader.GetInt32(12)
                val.Student = reader.GetInt32(13)
                returnval.Add(val)
            End While
            connection.Close()
            Return returnval
        End Function

        Public Shared Function delete(ByVal IdentifierQuery As String) As Integer
            Dim x As Integer
            Dim connection As New OleDbConnection(connnectionString)
            Dim command As New OleDbCommand("delete from PaymentTable where Identifier = @Identifier", connection)
            command.Parameters.Add("@Identifier", OleDbType.WChar).Value = IdentifierQuery
            connection.Open()
            x = command.ExecuteNonQuery
            connection.Close()
            Return x
        End Function
    End Class

    <Serializable()>
    Class QuarterGradingTable
        Implements IComparable(Of QuarterGradingTable)

        Public Property Identifier As String
        Public Property Quarter As String
        Public Property _Subject As String
        Public Property _Student As Integer
        Public Property Quiz As Integer
        Public Property Participation As Integer
        Public Property Recitation As Integer
        Public Property Project As Integer
        Public Property Exam As Integer
        Public Property FinalGrade As Integer
        Public Property SchoolYear As Integer

        Public Sub New()
        End Sub

        Public Function findStudentGrades(ByVal StudentID As Integer, Optional ByVal SchoolYear As Integer = 0, Optional ByVal Quarter As String = "") As List(Of QuarterGradingTable)
            Dim connection As New OleDbConnection(connnectionString)
            Dim command As OleDbCommand
            If SchoolYear = 0 And String.IsNullOrWhiteSpace(Quarter) Then
                command = New OleDbCommand("select * from QuarterGradingTable WHERE Student = @studentID", connection)
                command.Parameters.Add("@studentID", OleDbType.Integer).Value = StudentID
            ElseIf SchoolYear > 0 And String.IsNullOrWhiteSpace(Quarter) Then

            End If
        End Function
        Public Sub New(ByVal IdentifierQuery As String)
            Throw New Exception
        End Sub

        Public Sub save(ByVal Mode As SaveMode)
            If Mode = SaveMode.Insert Then
                insert()
            ElseIf Mode = SaveMode.Update Then
                update()
            End If
        End Sub

        Public Sub insert()
            Dim connection As New OleDbConnection(connnectionString)
            Dim command As New OleDbCommand("INSERT INTO QuarterGradingTable VALUES(@Identifier, @Quarter, @Subject, @Student,  @Quiz, @Participation, @Recitation, @Project, @Exam, @FinalGrade, @SchoolYear)", connection)
            command.Parameters.Add("@Identifier", OleDbType.WChar).Value = Identifier
            command.Parameters.Add("@Quarter", OleDbType.WChar).Value = Quarter
            command.Parameters.Add("@Subject", OleDbType.WChar).Value = _Subject
            command.Parameters.Add("@Student", OleDbType.Integer).Value = _Student
            command.Parameters.Add("@Quiz", OleDbType.Integer).Value = Quiz
            command.Parameters.Add("@Participation", OleDbType.Integer).Value = Participation
            command.Parameters.Add("@Recitation", OleDbType.Integer).Value = Recitation
            command.Parameters.Add("@Project", OleDbType.Integer).Value = Project
            command.Parameters.Add("@Exam", OleDbType.Integer).Value = Exam
            command.Parameters.Add("@FinalGrade", OleDbType.Integer).Value = FinalGrade
            command.Parameters.Add("@SchoolYear", OleDbType.Integer).Value = SchoolYear
            connection.Open()
            command.ExecuteNonQuery()
            connection.Close()
        End Sub

        Public Sub update()
            Dim connection As New OleDbConnection(connnectionString)
            Dim command As New OleDbCommand("UPDATE QuarterGradingTable SET Quarter = @Quarter, Subject = @Subject, Student = @Student, Quiz = @Quiz, Participation = @Participation, Recitation = @Recitation, Project = @Project, Exam = @Exam, FinalGrade = @FinalGrade, SchoolYear = @SchoolYear WHERE Identifier = @Identifier)", connection)
            command.Parameters.Add("@Quarter", OleDbType.WChar).Value = Quarter
            command.Parameters.Add("@Subject", OleDbType.WChar).Value = _Subject
            command.Parameters.Add("@Student", OleDbType.Integer).Value = _Student
            command.Parameters.Add("@Quiz", OleDbType.Integer).Value = Quiz
            command.Parameters.Add("@Participation", OleDbType.Integer).Value = Participation
            command.Parameters.Add("@Recitation", OleDbType.Integer).Value = Recitation
            command.Parameters.Add("@Project", OleDbType.Integer).Value = Project
            command.Parameters.Add("@Exam", OleDbType.Integer).Value = Exam
            command.Parameters.Add("@FinalGrade", OleDbType.Integer).Value = FinalGrade
            command.Parameters.Add("@SchoolYear", OleDbType.Integer).Value = SchoolYear
            command.Parameters.Add("@Identifier", OleDbType.WChar).Value = Identifier
            connection.Open()
            command.ExecuteNonQuery()
            connection.Close()
        End Sub

        Public Function getStudent() As Student
            Return Student.find(_Student)
        End Function

        Public Function getSubject() As Subject
            Return Subject.find(_Subject)
        End Function

        Public Shared Function find(ByVal IdentifierQuery As String) As QuarterGradingTable
            Dim returnValue As New QuarterGradingTable
            Dim connection As New OleDbConnection(connnectionString)
            Dim command As New OleDbCommand("SELECT * FROM QuarterGradingTable WHERE Identifier = @identifier", connection)
            Dim reader As OleDbDataReader
            command.Parameters.Add("@identifier", OleDbType.WChar).Value = IdentifierQuery
            connection.Open()
            reader = command.ExecuteReader
            While reader.Read
                returnValue.Identifier = reader.GetString(0)
                returnValue.Quarter = reader.GetString(1)
                returnValue._Subject = reader.GetString(2)
                returnValue._Student = reader.GetInt32(3)
                returnValue.Quiz = reader.GetInt32(4)
                returnValue.Participation = reader.GetInt32(5)
                returnValue.Recitation = reader.GetInt32(6)
                returnValue.Project = reader.GetInt32(7)
                returnValue.Exam = reader.GetInt32(8)
                returnValue.FinalGrade = reader.GetInt32(9)
                returnValue.SchoolYear = reader.GetInt32(10)
            End While
            connection.Close()
            Return returnValue
        End Function

        Public Shared Function findAll() As List(Of QuarterGradingTable)
            Dim curQ As QuarterGradingTable
            Dim returnValue As New List(Of QuarterGradingTable)
            Dim connection As New OleDbConnection(connnectionString)
            Dim command As New OleDbCommand("SELECT * FROM QuarterGradingTable", connection)
            Dim reader As OleDbDataReader
            connection.Open()
            reader = command.ExecuteReader
            While reader.Read
                curQ = New QuarterGradingTable
                curQ.Identifier = reader.GetString(0)
                curQ.Quarter = reader.GetString(1)
                curQ._Subject = reader.GetString(2)
                curQ._Student = reader.GetInt32(3)
                curQ.Quiz = reader.GetInt32(4)
                curQ.Participation = reader.GetInt32(5)
                curQ.Recitation = reader.GetInt32(6)
                curQ.Project = reader.GetInt32(7)
                curQ.Exam = reader.GetInt32(8)
                curQ.FinalGrade = reader.GetInt32(9)
                curQ.SchoolYear = reader.GetInt32(10)
                returnValue.Add(curQ)
            End While
            connection.Close()
            Return returnValue
        End Function

        Public Shared Function findBySchoolYear(ByVal year As Integer) As List(Of QuarterGradingTable)
            Dim curQ As QuarterGradingTable
            Dim returnValue As New List(Of QuarterGradingTable)
            Dim connection As New OleDbConnection(connnectionString)
            Dim command As New OleDbCommand("SELECT * FROM QuarterGradingTable where SchoolYear = @SY", connection)
            command.Parameters.Add("@SY", OleDbType.Integer).Value = year
            Dim reader As OleDbDataReader
            connection.Open()
            reader = command.ExecuteReader
            While reader.Read
                curQ = New QuarterGradingTable
                curQ.Identifier = reader.GetString(0)
                curQ.Quarter = reader.GetString(1)
                curQ._Subject = reader.GetString(2)
                curQ._Student = reader.GetInt32(3)
                curQ.Quiz = reader.GetInt32(4)
                curQ.Participation = reader.GetInt32(5)
                curQ.Recitation = reader.GetInt32(6)
                curQ.Project = reader.GetInt32(7)
                curQ.Exam = reader.GetInt32(8)
                curQ.FinalGrade = reader.GetInt32(9)
                curQ.SchoolYear = reader.GetInt32(10)
                returnValue.Add(curQ)
            End While
            connection.Close()
            Return returnValue
        End Function

        Public Shared Function delete(ByVal IdentifierQuery As String) As Integer
            Dim x As Integer
            Dim connection As New OleDbConnection(connnectionString)
            Dim command As New OleDbCommand("delete from QuarterGradingTable where Identifier = @Identifier", connection)
            command.Parameters.Add("@Identifier", OleDbType.WChar).Value = IdentifierQuery
            connection.Open()
            x = command.ExecuteNonQuery
            connection.Close()
            Return x
        End Function

        Public Function CompareTo(ByVal other As QuarterGradingTable) As Integer Implements System.IComparable(Of QuarterGradingTable).CompareTo
            Dim myStudent As Student = Student.find(_Student)
            Dim otherStudent As Student = Student.find(other._Student)
            Dim mySubject As Subject = Subject.find(_Subject)
            Dim otherSubject As Subject = Subject.find(other._Subject)
            If SchoolYear > other.SchoolYear Then
                Return 1
            ElseIf SchoolYear < other.SchoolYear Then
                Return -1
            ElseIf SchoolYear = other.SchoolYear Then
                If Quarter > other.Quarter Then
                    Return 1
                ElseIf Quarter < other.Quarter Then
                    Return -1
                ElseIf Quarter = other.Quarter Then
                    Dim x = mySubject.CompareTo(otherSubject)
                    If x > 0 Then
                        Return 1
                    ElseIf x < 0 Then
                        Return -1
                    ElseIf x = 0 Then
                        x = myStudent.CompareTo(otherStudent)
                        Return x
                    End If
                End If
            End If
            Return 0
        End Function
    End Class

    <Serializable()>
    Class Teacher
        Implements IComparable(Of Teacher)
        Public Property TeacherID As Integer
        Public Property LastName As String
        Public Property FirstName As String
        Public Property MiddleName As String
        Public Property Age As Integer
        Public Property Address As String
        Public Property Sex As String
        Public Property Birthday As Date
        Public Property SocialStatus As String
        Public Property Degree As String
        Public Property MajorSubject As String

        Public Sub New()
        End Sub

        Public Function getSubjects() As List(Of Subject)
            Dim returnVal As New List(Of Subject)
            Dim cursub As Subject
            Dim connection As New OleDbConnection(connnectionString)
            Dim command As New OleDbCommand("select * from Subjects where SubjectTeacher = @MyID", connection)
            command.Parameters.Add("@MyID", OleDbType.Integer).Value = TeacherID
            Dim reader As OleDbDataReader
            connection.Open()
            reader = command.ExecuteReader
            While reader.Read
                cursub = New Subject
                cursub.SubjectName = reader.GetString(0)
                cursub.SubjectTeacher = reader.GetInt32(1)
                cursub.Monday = reader.GetBoolean(2)
                cursub.MondayTime = reader.GetDateTime(3)
                cursub.Tuesday = reader.GetBoolean(4)
                cursub.TuesdayTime = reader.GetDateTime(5)
                cursub.Wednesday = reader.GetBoolean(6)
                cursub.WednesdayTime = reader.GetDateTime(7)
                cursub.Thursday = reader.GetBoolean(8)
                cursub.ThursdayTime = reader.GetDateTime(9)
                cursub.Friday = reader.GetBoolean(10)
                cursub.FridayTime = reader.GetDateTime(11)
                cursub.MondayTimeTO = reader.GetDateTime(12)
                cursub.TuesdayTimeTO = reader.GetDateTime(13)
                cursub.WednesdayTimeTO = reader.GetDateTime(14)
                cursub.ThursdayTimeTO = reader.GetDateTime(15)
                cursub.FridayTimeTO = reader.GetDateTime(16)
                returnVal.Add(cursub)
            End While
            reader.Close()
            connection.Close()
            Return returnVal
        End Function

        Public Function getAdvisoryClass() As Classes
            Dim returnval As New Classes
            Dim connection As New OleDbConnection(connnectionString)
            Dim command As New OleDbCommand("select * from Classes where ClassAdvisor = @ClassAdvisor", connection)
            Dim reader As OleDbDataReader
            command.Parameters.Add("@ClassAdvisor", OleDbType.Integer).Value = TeacherID
            connection.Open()
            reader = command.ExecuteReader
            While reader.Read
                returnval.ClassName = reader.GetString(0)
                returnval.ClassAdvisor = reader.GetInt32(1)
                returnval.ClassLevel = reader.GetInt32(2)
                returnval.NeedGraduation = reader.GetBoolean(3)
            End While
            If String.IsNullOrWhiteSpace(returnval.ClassName) Then
                Throw New Exception("Teacher doesn't have advisory class")
            End If
            reader.Close()
            connection.Close()
            Return returnval
        End Function

        Public Sub save(ByVal Mode As SaveMode)
            If Mode = SaveMode.Insert Then
                insert()
            ElseIf Mode = SaveMode.Update Then
                update()
            End If
        End Sub

        Public Sub insert()
            Dim connection As New OleDbConnection(connnectionString)
            Dim command As New OleDbCommand("INSERT INTO Teachers VALUES(@TeacherID, @LastName, @FirstName, @MiddleName, @Age, @Address, @Sex, @Birthday, @SocialStatus, @Degree, @MajorSubject)", connection)
            command.Parameters.Add("@TeacherID", OleDbType.Integer).Value = TeacherID
            command.Parameters.Add("@LastName", OleDbType.WChar).Value = LastName
            command.Parameters.Add("@FirstName", OleDbType.WChar).Value = FirstName
            command.Parameters.Add("@MiddleName", OleDbType.WChar).Value = MiddleName
            command.Parameters.Add("@Age", OleDbType.Integer).Value = Age
            command.Parameters.Add("@Address", OleDbType.WChar).Value = Address
            command.Parameters.Add("@Sex", OleDbType.WChar).Value = Sex
            command.Parameters.Add("@Birthday", OleDbType.DBDate).Value = Birthday
            command.Parameters.Add("@SocialStatus", OleDbType.WChar).Value = SocialStatus
            command.Parameters.Add("@Degree", OleDbType.WChar).Value = Degree
            command.Parameters.Add("@MajorSubject", OleDbType.WChar).Value = MajorSubject
            connection.Open()
            command.ExecuteNonQuery()
            connection.Close()
        End Sub

        Public Sub update()
            Dim connection As New OleDbConnection(connnectionString)
            Dim command As New OleDbCommand("UPDATE Teachers SET LastName = @LastName, FirstName = @FirstName, MiddleName = @MiddleName, Age = @Age, Address = @Address, Sex = @Sex, Birthday = @Birthday, SocialStatus = @SocialStatus, Degree = @Degree, MajorSubject = @MajorSubject WHERE TeacherID = @TeacherID", connection)
            command.Parameters.Add("@LastName", OleDbType.WChar).Value = LastName
            command.Parameters.Add("@FirstName", OleDbType.WChar).Value = FirstName
            command.Parameters.Add("@MiddleName", OleDbType.WChar).Value = MiddleName
            command.Parameters.Add("@Age", OleDbType.Integer).Value = Age
            command.Parameters.Add("@Address", OleDbType.WChar).Value = Address
            command.Parameters.Add("@Sex", OleDbType.WChar).Value = Sex
            command.Parameters.Add("@Birthday", OleDbType.DBDate).Value = Birthday
            command.Parameters.Add("@SocialStatus", OleDbType.WChar).Value = SocialStatus
            command.Parameters.Add("@Degree", OleDbType.WChar).Value = Degree
            command.Parameters.Add("@MajorSubject", OleDbType.WChar).Value = MajorSubject
            command.Parameters.Add("@TeacherID", OleDbType.Integer).Value = TeacherID
            connection.Open()
            command.ExecuteNonQuery()
            connection.Close()
        End Sub

        Public Shared Function find(ByVal TeacherIDQuery As Integer) As Teacher
            Dim returnValue As New Teacher
            Dim connection As New OleDbConnection(connnectionString)
            Dim command As New OleDbCommand("SELECT * FROM Teachers WHERE TeacherID = @TeacherID", connection)
            Dim reader As OleDbDataReader
            command.Parameters.Add("@TeacherID", OleDbType.Integer).Value = TeacherIDQuery
            connection.Open()
            reader = command.ExecuteReader
            While reader.Read
                returnValue.TeacherID = reader.GetInt32(0)
                returnValue.LastName = reader.GetString(1)
                returnValue.FirstName = reader.GetString(2)
                returnValue.MiddleName = reader.GetString(3)
                returnValue.Age = reader.GetInt32(4)
                returnValue.Address = reader.GetString(5)
                returnValue.Sex = reader.GetString(6)
                returnValue.Birthday = reader.GetDateTime(7)
                returnValue.SocialStatus = reader.GetString(8)
                returnValue.Degree = reader.GetString(9)
                returnValue.MajorSubject = reader.GetString(10)
            End While
            connection.Close()
            Return returnValue
        End Function

        Public Shared Function findAll() As List(Of Teacher)
            Dim returnValue As New List(Of Teacher)
            Dim curTeacher As Teacher
            Dim connection As New OleDbConnection(connnectionString)
            Dim command As New OleDbCommand("SELECT * FROM Teachers", connection)
            Dim reader As OleDbDataReader
            connection.Open()
            reader = command.ExecuteReader
            While reader.Read
                curTeacher = New Teacher
                curTeacher.TeacherID = reader.GetInt32(0)
                curTeacher.LastName = reader.GetString(1)
                curTeacher.FirstName = reader.GetString(2)
                curTeacher.MiddleName = reader.GetString(3)
                curTeacher.Age = reader.GetInt32(4)
                curTeacher.Address = reader.GetString(5)
                curTeacher.Sex = reader.GetString(6)
                curTeacher.Birthday = reader.GetDateTime(7)
                curTeacher.SocialStatus = reader.GetString(8)
                curTeacher.Degree = reader.GetString(9)
                curTeacher.MajorSubject = reader.GetString(10)
                returnValue.Add(curTeacher)
            End While
            connection.Close()
            Return returnValue
        End Function

        Public Shared Function delete(ByVal TeacherIDuUery As Integer) As Integer
            Dim x As Integer
            Dim connection As New OleDbConnection(connnectionString)
            Dim command As New OleDbCommand("delete from Teacher where TeacherID = @Identifier", connection)
            command.Parameters.Add("@Identifier", OleDbType.Integer).Value = TeacherIDuUery
            connection.Open()
            x = command.ExecuteNonQuery
            connection.Close()
            Return x
        End Function

        Public Function CompareTo(ByVal other As Teacher) As Integer Implements System.IComparable(Of Teacher).CompareTo
            If LastName = other.LastName Then
                If FirstName = other.FirstName Then
                    If MiddleName = other.MiddleName Then
                        Return 0
                    ElseIf MiddleName > other.MiddleName Then
                        Return 1
                    ElseIf MiddleName < other.MiddleName Then
                        Return -1
                    End If
                ElseIf FirstName > other.FirstName Then
                    Return 1
                ElseIf FirstName < other.FirstName Then
                    Return -1
                End If
            ElseIf LastName > other.LastName Then
                Return 1
            ElseIf LastName < other.LastName Then
                Return -1
            End If
            Return 0
        End Function

        Public Overrides Function ToString() As String
            Return FirstName & " " & MiddleName(0) & ". " & LastName
        End Function
    End Class

    ''' <summary>
    ''' A Subject object that is connected to a database
    ''' </summary>
    ''' <remarks>Contains the scheduled time of itself, it's name and it's subject teacher and the class it is taught.</remarks>
    <Serializable()>
    Class Subject
        Implements IComparable(Of Subject)
        Public Property SubjectName As String
        Public Property SubjectTeacher As Integer
        Public Property Monday As Boolean
        Public Property MondayTime As Date
        Public Property Tuesday As Boolean
        Public Property TuesdayTime As Date
        Public Property Wednesday As Boolean
        Public Property WednesdayTime As Date
        Public Property Thursday As Boolean
        Public Property ThursdayTime As Date
        Public Property Friday As Boolean
        Public Property FridayTime As Date
        Public Property MondayTimeTO As Date
        Public Property TuesdayTimeTO As Date
        Public Property WednesdayTimeTO As Date
        Public Property ThursdayTimeTO As Date
        Public Property FridayTimeTO As Date

        Public Sub New()
        End Sub

        Public Overrides Function ToString() As String
            Return SubjectName
        End Function

        Public Function getTeacher() As Teacher
            Return Teacher.find(SubjectTeacher)
        End Function

        Public Function getSubjectGrades() As List(Of QuarterGradingTable)
            Dim curQ As QuarterGradingTable
            Dim returnValue As New List(Of QuarterGradingTable)
            Dim connection As New OleDbConnection(connnectionString)
            Dim command As New OleDbCommand("SELECT * FROM QuarterGradingTable WHERE Subject = @MyName", connection)
            command.Parameters.Add("@MyName", OleDbType.WChar).Value = SubjectName
            Dim reader As OleDbDataReader
            connection.Open()
            reader = command.ExecuteReader
            While reader.Read
                curQ = New QuarterGradingTable
                curQ.Identifier = reader.GetString(0)
                curQ.Quarter = reader.GetString(1)
                curQ._Subject = reader.GetString(2)
                curQ._Student = reader.GetInt32(3)
                curQ.Quiz = reader.GetInt32(4)
                curQ.Participation = reader.GetInt32(5)
                curQ.Recitation = reader.GetInt32(6)
                curQ.Project = reader.GetInt32(7)
                curQ.Exam = reader.GetInt32(8)
                curQ.FinalGrade = reader.GetInt32(9)
                curQ.SchoolYear = reader.GetInt32(10)
                returnValue.Add(curQ)
            End While
            reader.Close()
            connection.Close()
            Return returnValue
        End Function

        Public Sub save(ByVal mode As SaveMode)
            If mode = SaveMode.Insert Then
                insert()
            ElseIf mode = SaveMode.Update Then
                update()
            End If
        End Sub

        Public Sub insert()
            Dim connection As New OleDbConnection(connnectionString)
            Dim command As New OleDbCommand("INSERT INTO Subjects VALUES(@SubjectName, @SubjectTeacher, @Monday, @MondayTime, @Tuesday, @TuesdayTime, @Wednesday, @WednesdayTime, @Thursday, @ThursdayTime, @Friday, @FridayTime, @MondayTimeTo, @TuesdayTimeTo, @WednesdayTimeTo, @ThursdayTimeTo, @FridayTimeTo)", connection)
            command.Parameters.Add("@SubjectName", OleDbType.WChar).Value = SubjectName
            command.Parameters.Add("@SubjectTeacher", OleDbType.Integer).Value = SubjectTeacher
            command.Parameters.Add("@Monday", OleDbType.Boolean).Value = Monday
            command.Parameters.Add("@MondayTime", OleDbType.DBTime).Value = MondayTime
            command.Parameters.Add("@Tuesday", OleDbType.Boolean).Value = Tuesday
            command.Parameters.Add("@TuesdayTime", OleDbType.DBTime).Value = TuesdayTime
            command.Parameters.Add("@Wednesday", OleDbType.Boolean).Value = Wednesday
            command.Parameters.Add("@WednesdayTime", OleDbType.DBTime).Value = WednesdayTime
            command.Parameters.Add("@Thursday", OleDbType.Boolean).Value = Thursday
            command.Parameters.Add("@ThursdayTime", OleDbType.DBTime).Value = ThursdayTime
            command.Parameters.Add("@Friday", OleDbType.Boolean).Value = Friday
            command.Parameters.Add("@FridayTime", OleDbType.DBTime).Value = FridayTime
            command.Parameters.Add("@MondayTimeTo", OleDbType.DBTime).Value = MondayTimeTO
            command.Parameters.Add("@TuesdayTimeTo", OleDbType.DBTime).Value = TuesdayTimeTO
            command.Parameters.Add("@WednesdayTimeTo", OleDbType.DBTime).Value = WednesdayTimeTO
            command.Parameters.Add("@ThursdayTimeTo", OleDbType.DBTime).Value = ThursdayTimeTO
            command.Parameters.Add("@FridayTimeTo", OleDbType.DBTime).Value = FridayTimeTO
            connection.Open()
            command.ExecuteNonQuery()
            connection.Close()
        End Sub

        Public Sub update()
            Dim connection As New OleDbConnection(connnectionString)
            Dim command As New OleDbCommand("Update Subjects SET SubjectTeacher = @SubjectTeacher, Monday = @Monday, MondayTime = @MondayTime, Tuesday = @Tuesday, TuesdayTime = @TuesdayTime, Wednesday = @Wednesday, WednesdayTIme = @WednesdayTime, Thursday = @Thursday, ThursdayTime = @ThursdayTime, Friday = @Friday, FridayTime = @FridayTime, MondayTimeTo = @MondayTimeTo, TuesdayTimeTo = @TuesdayTimeTo, WednesdayTimeTo = @WednesdayTimeTo, ThursdayTimeTo = @ThursdayTimeTo, FridayTimeTo = @FridayTimeTo WHERE SubjectName = @SubjectName", connnectionString)
            command.Parameters.Add("@SubjectTeacher", OleDbType.Integer).Value = SubjectTeacher
            command.Parameters.Add("@Monday", OleDbType.Boolean).Value = Monday
            command.Parameters.Add("@MondayTime", OleDbType.DBTime).Value = MondayTime
            command.Parameters.Add("@Tuesday", OleDbType.Boolean).Value = Tuesday
            command.Parameters.Add("@TuesdayTime", OleDbType.DBTime).Value = TuesdayTime
            command.Parameters.Add("@Wednesday", OleDbType.Boolean).Value = Wednesday
            command.Parameters.Add("@WednesdayTime", OleDbType.DBTime).Value = WednesdayTime
            command.Parameters.Add("@Thursday", OleDbType.Boolean).Value = Thursday
            command.Parameters.Add("@ThursdayTime", OleDbType.DBTime).Value = ThursdayTime
            command.Parameters.Add("@Friday", OleDbType.Boolean).Value = Friday
            command.Parameters.Add("@FridayTime", OleDbType.DBTime).Value = FridayTime
            command.Parameters.Add("@MondayTimeTo", OleDbType.DBTime).Value = MondayTimeTO
            command.Parameters.Add("@TuesdayTimeTo", OleDbType.DBTime).Value = TuesdayTimeTO
            command.Parameters.Add("@WednesdayTimeTo", OleDbType.DBTime).Value = WednesdayTimeTO
            command.Parameters.Add("@ThursdayTimeTo", OleDbType.DBTime).Value = ThursdayTimeTO
            command.Parameters.Add("@FridayTimeTo", OleDbType.DBTime).Value = FridayTimeTO
            command.Parameters.Add("@SubjectName", OleDbType.WChar).Value = SubjectName
            connection.Open()
            command.ExecuteNonQuery()
            connection.Close()
        End Sub

        Public Shared Function find(ByVal SubjectNameQuery As String) As Subject
            Dim returnVal As New Subject
            Dim connection As New OleDbConnection(connnectionString)
            Dim command As New OleDbCommand("select * from Subjects where SubjectName = @SubjectName", connection)
            Dim reader As OleDbDataReader
            command.Parameters.Add("@SubjectName", OleDbType.WChar).Value = SubjectNameQuery
            connection.Open()
            reader = command.ExecuteReader
            While reader.Read
                returnVal.SubjectName = reader.GetString(0)
                returnVal.SubjectTeacher = reader.GetInt32(1)
                returnVal.Monday = reader.GetBoolean(2)
                returnVal.MondayTime = reader.GetDateTime(3)
                returnVal.Tuesday = reader.GetBoolean(4)
                returnVal.TuesdayTime = reader.GetDateTime(5)
                returnVal.Wednesday = reader.GetBoolean(6)
                returnVal.WednesdayTime = reader.GetDateTime(7)
                returnVal.Thursday = reader.GetBoolean(8)
                returnVal.ThursdayTime = reader.GetDateTime(9)
                returnVal.Friday = reader.GetBoolean(10)
                returnVal.FridayTime = reader.GetDateTime(11)
                returnVal.MondayTimeTO = reader.GetDateTime(12)
                returnVal.TuesdayTimeTO = reader.GetDateTime(13)
                returnVal.WednesdayTimeTO = reader.GetDateTime(14)
                returnVal.ThursdayTimeTO = reader.GetDateTime(15)
                returnVal.FridayTimeTO = reader.GetDateTime(16)
            End While
            reader.Close()
            connection.Close()
            Return returnVal
        End Function

        Public Shared Function findall() As List(Of Subject)
            Dim returnVal As New List(Of Subject)
            Dim cursub As Subject
            Dim connection As New OleDbConnection(connnectionString)
            Dim command As New OleDbCommand("select * from Subjects", connection)
            Dim reader As OleDbDataReader
            connection.Open()
            reader = command.ExecuteReader
            While reader.Read
                cursub = New Subject
                cursub.SubjectName = reader.GetString(0)
                cursub.SubjectTeacher = reader.GetInt32(1)
                cursub.Monday = reader.GetBoolean(2)
                cursub.MondayTime = reader.GetDateTime(3)
                cursub.Tuesday = reader.GetBoolean(4)
                cursub.TuesdayTime = reader.GetDateTime(5)
                cursub.Wednesday = reader.GetBoolean(6)
                cursub.WednesdayTime = reader.GetDateTime(7)
                cursub.Thursday = reader.GetBoolean(8)
                cursub.ThursdayTime = reader.GetDateTime(9)
                cursub.Friday = reader.GetBoolean(10)
                cursub.FridayTime = reader.GetDateTime(11)
                cursub.MondayTimeTO = reader.GetDateTime(12)
                cursub.TuesdayTimeTO = reader.GetDateTime(13)
                cursub.WednesdayTimeTO = reader.GetDateTime(14)
                cursub.ThursdayTimeTO = reader.GetDateTime(15)
                cursub.FridayTimeTO = reader.GetDateTime(16)
                returnVal.Add(cursub)
            End While
            reader.Close()
            connection.Close()
            Return returnVal
        End Function

        Public Shared Function delete(ByVal SubjectNameQuery As String) As Integer
            Dim x As Integer
            Dim connection As New OleDbConnection(connnectionString)
            Dim command As New OleDbCommand("delete from Subjects where SubjectName = @Identifier", connection)
            command.Parameters.Add("@Identifier", OleDbType.WChar).Value = SubjectNameQuery
            connection.Open()
            x = command.ExecuteNonQuery
            connection.Close()
            Return x
        End Function

        Public Shared Function GenerateHTMLTable() As IO.Path
            Throw New Exception("Unimplemented Exception")
        End Function

        Public Function CompareTo(ByVal other As Subject) As Integer Implements System.IComparable(Of Subject).CompareTo

        End Function
    End Class

    ''' <summary>
    ''' A class for the database table Constant.
    ''' Use Constant.fetch(ConstantName) to get constants from the database.
    ''' </summary>
    ''' <remarks>Implemented</remarks>
    <Serializable()>
    Class Constants

        ''' <summary>
        ''' A method for getting a constant from the database.
        ''' If you provided a wrong constant name then nothing will happen.
        ''' Make sure you provide the correct ConstantName
        ''' </summary>
        ''' <param name="ConstantName">The name of the constant from the table.</param>
        ''' <returns>Integer</returns>
        ''' <remarks>Implemented</remarks>
        Public Shared Function Fetch(ByVal ConstantName As String) As Int32
            Dim x As Integer
            Dim connection As New OleDbConnection(connnectionString)
            Dim command As New OleDbCommand("Select ConstantValue from Constants where ConstantName = @ConstantName", connection)
            command.Parameters.Add("@ConstantName", OleDbType.Integer).Value = ConstantName
            connection.Open()
            x = command.ExecuteScalar
            connection.Close()
            Return x
        End Function

        ''' <summary>
        ''' Method on making a constant to the table.
        ''' Make sure you provide a unique ConstantName or it will raise an error
        ''' </summary>
        ''' <param name="ConstantName">Name of the constant, must be unique to the table</param>
        ''' <param name="ConstantValue">Value of the constant</param>
        ''' <remarks>Implemented</remarks>
        Public Shared Sub Make(ByVal ConstantName As String, ByVal ConstantValue As Integer)
            Dim connection As New OleDbConnection(connnectionString)
            Dim command As New OleDbCommand("INSERT INTO Constants VALUES(@Name, @Value)", connection)
            Dim validationResult As Integer
            command.Parameters.Add("@Name", OleDbType.WChar).Value = ConstantName
            command.Parameters.Add("@Value", OleDbType.WChar).Value = ConstantValue
            Dim validationCommand As New OleDbCommand("select * from Constants where ConstantName = @name", connection)
            command.Parameters.Add("@Name", OleDbType.WChar).Value = ConstantName
            connection.Open()
            validationResult = validationCommand.ExecuteNonQuery
            connection.Close()
            If validationResult > 0 Then
                Throw New Exception("Constant already exist. Why not update it instead")
            End If
            connection.Open()
            command.ExecuteNonQuery()
            connection.Close()
        End Sub

        ''' <summary>
        ''' Method of changing constants to the table.
        ''' Make sure you provide an existing ConstantName or it will raise an error.
        ''' </summary>
        ''' <param name="ConstantName">Name of the Constant, Must be existing</param>
        ''' <param name="ConstantValue">Value of the constant</param>
        ''' <remarks>Implemented</remarks>
        Public Shared Sub Update(ByVal ConstantName As String, ByVal ConstantValue As Integer)
            Dim connection As New OleDbConnection(connnectionString)
            Dim command As New OleDbCommand("UPDATE Constants SET ConstantValue = @Value Where ConstantName = @Name", connection)
            Dim validationResult As Integer
            command.Parameters.Add("@Value", OleDbType.WChar).Value = ConstantValue
            command.Parameters.Add("@Name", OleDbType.WChar).Value = ConstantName
            Dim validationCommand As New OleDbCommand("select * from Constants where ConstantName = @name", connection)
            command.Parameters.Add("@Name", OleDbType.WChar).Value = ConstantName
            connection.Open()
            validationResult = validationCommand.ExecuteNonQuery
            connection.Close()
            If validationResult < 0 Then
                Throw New Exception("Constant does not exist. Why not making it instead")
            End If
            connection.Open()
            command.ExecuteNonQuery()
            connection.Close()
        End Sub

        ''' <summary>
        ''' A function for deletion of constants. Returns 1 if the constant is succesfully deleted else returns 0
        ''' </summary>
        ''' <param name="constantName">Name of the constant to be deleted</param>
        ''' <returns>Integer</returns>
        ''' <remarks>Implemented</remarks>
        Public Shared Function Delete(ByVal constantName As String) As Integer
            Dim x As Integer
            Dim connection As New OleDbConnection(connnectionString)
            Dim command As New OleDbCommand("delete from Constants where ConstantName = @Identifier", connection)
            command.Parameters.Add("@Identifier", OleDbType.WChar).Value = constantName
            connection.Open()
            x = command.ExecuteNonQuery
            connection.Close()
            Return x
        End Function
    End Class

    ''' <summary>
    ''' </summary>
    ''' <remarks></remarks>
    <Serializable()>
    Class User
        Public Shared Function Login(ByVal username As String, ByVal password As String) As Boolean
            Dim returnValue = False
            Dim dbuser As String = ""
            Dim dbpass As String = ""
            Dim dbtype As String = ""
            Dim connnection As New OleDbConnection(connnectionString)
            Dim command As New OleDbCommand("select * from Users where Username = @username AND Userpassword = @password", connnection)
            command.Parameters.Add("@username", OleDbType.WChar).Value = username
            command.Parameters.Add("@password", OleDbType.WChar).Value = password
            Dim reader As OleDbDataReader
            connnection.Open()
            reader = command.ExecuteReader
            While reader.Read
                dbuser = reader.GetString(0)
                dbpass = reader.GetString(1)
                dbtype = reader.GetString(2)
                If dbtype = "Teacher" Then
                    TopLevelProperties.isTeacher = True
                    TopLevelProperties.TeacherID = reader.GetInt32(3)
                End If
            End While
            reader.Close()
            connnection.Close()
            If dbuser = username And dbpass = password Then
                returnValue = True
                TopLevelProperties.UserName = dbuser
                TopLevelProperties.UserType = dbtype
            End If
            Return returnValue
        End Function

        Public Shared Sub createUser(ByVal username As String, ByVal userpass As String, ByVal usertype As String, Optional ByVal teacherID As Integer = 0)
            Dim connection As New OleDbConnection(connnectionString)
            If usertype = "Teacher" Then
                Dim command As New OleDbCommand("insert into Users VALUES(@username, @userpass, @usertype, @teacherID)")
                command.Parameters.Add("@username", OleDbType.WChar).Value = username
                command.Parameters.Add("@userpass", OleDbType.WChar).Value = userpass
                command.Parameters.Add("@usertype", OleDbType.WChar).Value = usertype
                command.Parameters.Add("@teacherID", OleDbType.Integer).Value = teacherID
                connection.Open()
                command.ExecuteNonQuery()
                connection.Close()
            Else
                Dim command As New OleDbCommand("insert into Users(Username, Userpassword, UserType) VALUES(@username, @userpass, @usertype)")
                command.Parameters.Add("@username", OleDbType.WChar).Value = username
                command.Parameters.Add("@userpass", OleDbType.WChar).Value = userpass
                command.Parameters.Add("@usertype", OleDbType.WChar).Value = usertype
                connection.Open()
                command.ExecuteNonQuery()
                connection.Close()
            End If
        End Sub

        Public Shared Sub logout()
            TopLevelProperties.isTeacher = False
            TopLevelProperties.TeacherID = 0
            TopLevelProperties.UserName = vbNullString
            TopLevelProperties.UserType = vbNullString
        End Sub

        Public Shared Function changeUsernameAndPassword(ByVal username As String, ByVal userpass As String, ByVal newusername As String, ByVal newpass As String) As Integer
            Dim connection As New OleDbConnection(connnectionString)
            Dim command As New OleDbCommand("UPDATE Users set Username = @newusername, UserPassword = @newpass where Username = @oldusername AND Userpassword = @olduserpass", connection)
            command.Parameters.Add("@newusername", OleDbType.WChar).Value = newusername
            command.Parameters.Add("@newpass", OleDbType.WChar).Value = newpass
            command.Parameters.Add("@oldusername", OleDbType.WChar).Value = username
            command.Parameters.Add("@olduserpass", OleDbType.WChar).Value = userpass
            connection.Open()
            Dim x = command.ExecuteNonQuery
            connection.Close()
            Return x
        End Function
    End Class

    ''' <summary>
    ''' Since Subjects and Classes have a many-to-many relationship.
    ''' They need to have a seperate table for their relationship.
    ''' Thus also needing a seperate class.
    ''' </summary>
    ''' <remarks>Not Fully Implemented Yet</remarks>
    Class SubjectClass
        ''' <summary>
        ''' A method for connecting a class to a subject therefore saying this class teaches this subject.
        ''' </summary>
        ''' <param name="myClasss">The class to be connected</param>
        ''' <param name="subject">To the subject</param>
        ''' <remarks>Implemented</remarks>
        Public Shared Sub connect(ByVal myClasss As Classes, ByVal subject As Subject)
            Dim connection As New OleDbConnection(connnectionString)
            Dim command As New OleDbCommand("INSERT INTO SubjectClass VALUES(@class, @subject)", connection)
            command.Parameters.Add("@class", OleDbType.WChar).Value = myClasss.ClassName
            command.Parameters.Add("@subject", OleDbType.WChar).Value = subject.SubjectName
            connection.Open()
            command.ExecuteNonQuery()
            connection.Close()
        End Sub

        ''' <summary>
        ''' A method for getting the subjects all taught in one class. 
        ''' </summary>
        ''' <param name="ClassName">The name of the class you want to get their subjects on</param>
        ''' <returns>List(Of Subject)</returns>
        ''' <remarks>Implemented</remarks>
        Public Shared Function getClassSubjects(ByVal ClassName As String) As List(Of Subject)
            Dim returnVal As New List(Of Subject)
            Dim listOfSubjectName As New List(Of String)
            Dim connection As New OleDbConnection(connnectionString)
            Dim command As New OleDbCommand("select Subject from SubjectClass where Class = @ClassName", connection)
            Dim reader As OleDbDataReader
            command.Parameters.Add("@ClassName", OleDbType.WChar).Value = ClassName
            connection.Open()
            reader = command.ExecuteReader
            While reader.Read
                listOfSubjectName.Add(reader.GetString(0))
            End While
            reader.Close()
            connection.Close()
            For Each element In listOfSubjectName
                returnVal.Add(Subject.find(element))
            Next
            Return returnVal
        End Function

        ''' <summary>
        ''' To disconnect or to tell a class that they are not teaching that kind of subject anymore
        ''' returns 1 if disconnect is successful else 0
        ''' </summary>
        ''' <param name="myClasss">The classname of the class you want to disconnect</param>
        ''' <param name="subject">The name of the subject</param>
        ''' <returns>Integer</returns>
        ''' <remarks>Implemented</remarks>
        Public Shared Function disconnect(ByVal myClasss As String, ByVal subject As String) As Integer
            Dim connection As New OleDbConnection(connnectionString)
            Dim command As New OleDbCommand("Delete from SubjectClass where Class = @myClass AND Subject = @subject", connection)
            command.Parameters.Add("@myClass", OleDbType.WChar).Value = myClasss
            command.Parameters.Add("@subject", OleDbType.WChar).Value = subject
            connection.Open()
            Dim x As Integer = command.ExecuteNonQuery
            connection.Close()
            Return x
        End Function
    End Class
End Module