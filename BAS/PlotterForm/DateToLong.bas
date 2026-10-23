Private Function DateToLong(myDate As Date) As Long
      DateToLong = Year(myDate) * 10000 + Month(myDate) * 100 + Day(myDate)
End Function
