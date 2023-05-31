' Detail class:
' Used to construct the detail into an object
'
' Detail will contain properties:
'    ~Detail Number
'    ~Cost Value
'    ~CC (Cost Code) value
'    ~Rate value
'    ~Hours value
'    ~Description value
'    ~Total value
'    ~Margin Total value
'    ~Detail list (Collection of all detail objects)
'--------------------------------------------------------------------------------

' Initialize all variables for properties
Private detailNo As Integer
Private costVal As Integer
Private ccVal As String
Private rateVal As Double
Private hoursVal As Double
Private descriptionVal As String
Private totalVal As Double
Private marginTotalVal As Double
Public detail_list As Collection

' Detail constructor
' Set values to a temporary value. Replaced in function when each detail is created
Private Sub Class_Initialize()

    detailNo = 1
    costVal = 1
    ccVal = "test"
    rateVal = 1#
    hoursVal = 1#
    descriptionVal = "test"
    totalVal = 1#
    marginTotalVal = 1#
    
End Sub

'intitalize the collection in the constructor of the class
'Private Sub Class_Initialize()
'    Set detail_list = New Collection
'End Sub


' Getter and setter properties
Public Property Let Set_Detail(vDetail As Integer)
    detailNo = vDetail
End Property

Public Property Get detail() As Integer
    detail = detailNo
End Property

Public Property Let Set_Cost(vcost As Integer)
    costVal = vcost
End Property

Public Property Get Cost() As Integer
    Cost = costVal
End Property

Public Property Let Set_CC(vcc As String)
    ccVal = vcc
End Property

Public Property Get CC() As String
    CC = ccVal
End Property

Public Property Let Set_Rate(vrate As Double)
    rateVal = vrate
End Property

Public Property Get Rate() As Double
    Rate = rateVal
End Property

Public Property Let Set_Hours(vhours As Double)
    hoursVal = vhours
End Property

Public Property Get Hours() As Double
    Hours = hoursVal
End Property

Public Property Let Set_Description(vdescription As String)
    descriptionVal = vdescription
End Property

Public Property Get Description() As String
    Description = descriptionVal
End Property

Public Property Let Set_Total(vtotal As Double)
    totalVal = vtotal
End Property

Public Property Get Total() As Double
    Total = totalVal
End Property

Public Property Let Set_marginTotal(vmargintotal As Double)
    marginTotalVal = vmargintotal
End Property

Public Property Get marginTotal() As Double
    marginTotal = marginTotalVal
End Property


