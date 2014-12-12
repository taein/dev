<%@ CodePage = "65001" %>
<% 
	Option explicit
	Response.CharSet = "utf-8"
	Response.ContentType="application/javascript"
	
	Dim strDate, loopi, dtStDate, dtEndDate
	strDate = request("date")
	If len(strDate) = 6 Then
		dtStDate = left(strDate,4) & "-" & right(strDate,2) & "-01"
		dtStDate = cdate(dtStDate)
		dtEndDate = dateAdd("d",-1,dateAdd("m",1,dtStDate))
	
		With Response
			.Write "{"
			.Write """nowDate"":""" & left(strDate,4) & "년 " & right(strDate,2) & "월"","
			.Write """weekDay"":" & weekday(dtStDate) & ","
			.Write """calendar"":["
			For loopi = 1 To dateDiff("d",dtStDate,dtEndDate)+1
				.Write loopi
				If loopi <> dateDiff("d",dtStDate,dtEndDate)+1 Then
					.Write ","
				End If
			Next
			.Write "]"
			.Write "}"	
		End With
	End If
	
%>

