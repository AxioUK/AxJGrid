Attribute VB_Name = "oTest"

'Simular db
Public Enum NameTypeEnum
    ntRandom = 0
    ntMale = 1
    ntFemale = 2
End Enum

#If False Then
    Private ntRandom, ntMale, ntFemale
#End If

'These are for generating the demo data
Private Const M_FORENAMES = "Alan,Alfie,Andrew,Ben,Bill,Bob,Boris,Brian,Charles,Chris,David,Gavin,Geoff,Grant,Harry,Ian,James,Jon,Mark,Matthew,Michael,Patrick,Paul,Peter,Richard,Robert,Samuel,Simon,Tony,Trevor,William"
Private Const F_FORENAMES = "Alicia,Alison,Amanda,Barbara,Caroline,Charlotte,Dawn,Hannah,Harriet,Hayley,Jane,Jennifer,Karen,Katie,Kerry,Kim,Lara,Laura,Lucy,Mary,Mellisa,Patricia,Paula,Rachel,Sarah,Stephanie,Susan,Tracy,Vanessa"
Private Const SURNAMES = "Anderson-Allen,Black Evans,Bloggs,Brown,Clarke,Cole,Davis,Dawson Gate,Evans Brown,Gate ,Johnson Gate,Jones,Lawson,Lee,Richards,Ryan,Smith,Stephens,Temple,Turner,Wallace,White,Williams"
   
Private Const JOBS = "Accountant,Architect,Artist,Banker,Builder,Carpenter,Dentist,Director,Doctor,Engineer,Estate Agent,Fire Fighter,Gardener,Manager,Mechanic,Miner,Nurse,Optician,Pilot,Plumber,Police,Programmer,Scientist,Secretary,Shop Assistant,Solicitor,Surgeon,Teacher,Truck Driver,Vet"
   
Private mCalled As Boolean

Private mMF() As String
Private mFF() As String
Private mSurnames() As String

Private mJobs() As String
 
Public Function GetJobName(Optional Index As Long = -1) As String
    Initialise
    
    If Index = -1 Then
        GetJobName = mJobs(RandomInt(LBound(mJobs), UBound(mJobs)))
    Else
        GetJobName = mJobs(Index)
    End If
End Function

Public Function GetSurname() As String
    Initialise
    
    GetSurname = mSurnames(RandomInt(LBound(mSurnames), UBound(mSurnames)))
End Function


Public Function GetForename(Optional nType As NameTypeEnum) As String
    Initialise
    
    Select Case nType
        Case ntRandom
            If RandomInt(0, 1) = 0 Then
                GetForename = mMF(RandomInt(LBound(mMF), UBound(mMF)))
            Else
                GetForename = mFF(RandomInt(LBound(mFF), UBound(mFF)))
            End If
           
        Case ntMale
            GetForename = mMF(RandomInt(LBound(mMF), UBound(mMF)))

        Case ntFemale
            GetForename = mFF(RandomInt(LBound(mFF), UBound(mFF)))

    End Select
End Function



Public Function GetNameOfPerson(Optional nType As NameTypeEnum) As String
    Select Case nType
        Case ntRandom
            If RandomInt(0, 1) = 0 Then
                GetNameOfPerson = GetForename(ntMale) & " " & GetSurname()
            Else
                GetNameOfPerson = GetForename(ntFemale) & " " & GetSurname()
            End If
           
        Case ntMale
            GetNameOfPerson = GetForename(ntMale) & " " & GetSurname()

        Case ntFemale
            GetNameOfPerson = GetForename(ntFemale) & " " & GetSurname()

    End Select
End Function



Private Sub Initialise()
    If Not mCalled Then
        mCalled = True
        Randomize Timer
        
        mMF() = Split(M_FORENAMES, ",")
        mFF() = Split(F_FORENAMES, ",")
        mSurnames() = Split(SURNAMES, ",")
        
        mJobs() = Split(JOBS, ",")
    End If
End Sub

Public Function JobCount() As Long
    Initialise
    JobCount = UBound(mJobs)
End Function

Public Function RandomInt(lowerbound As Long, upperbound As Long) As Long
    RandomInt = Int((upperbound - lowerbound + 1) * Rnd + lowerbound)
End Function




