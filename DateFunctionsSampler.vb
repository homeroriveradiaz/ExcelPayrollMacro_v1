
Function Todays()
    Todays = Date
End Function

Function TodayWithTime()
    TodayWithTime = Now
End Function

Function AddHours(Hours, MyDate)
    AddHours = DateAdd("h", Hours, MyDate)
End Function

Function GetFirstDateOfTheMonth(MyDate)
	'Day(dateValue) returns the day of the month (e.g. 1 - 31)
	'You can determine the first day of a month by passing any given date within that month 
	'Pretty useful when working with monthly intervals
    GetFirstDateOfTheMonth = DateAdd("d", -Day(MyDate) + 1, MyDate)
End Function

Function HowManyMonthsApart(FirstDate, SecondDate)
    'Earliest date goes first
	'Which means you'll get negatives otherwise
	'Maybe that's what you want, who knows
    HowManyMonthsApart = DateDiff("m", FirstDate, SecondDate)
End Function

