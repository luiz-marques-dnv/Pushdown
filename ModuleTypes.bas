Type StepRecord
    StepNo As Integer
    TimeInSeconds As Double
    AccumulatedEnergy As Double
    MaxHeatInput As Double
    MaxTemp As Double
    MinTemp As Double
End Type

' Declare a custom type (similar to a Python dictionary) to store the data
Type SimulationData
    TotalSimTime As Double
    NumberOfTimeSteps As Double
    BeltempTimeInterval As Double
    StepData() As StepRecord
End Type


Type elementType
    Number As Long
    Strain As Double

End Type
