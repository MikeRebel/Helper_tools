Attribute VB_Name = "Functions"
Public Function Weighted_Mean(MyMultSum, MySumm)

x = 0
y = 0
Z = 0

For i = 1 To MyMultSum.Count
    x = MyMultSum(i) * MySumm(i)
    y = y + x
    Z = Z + MySumm(i)
Next i

Weighted_Mean = y / Z

End Function

