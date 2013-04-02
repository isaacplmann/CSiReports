<%
function YesNoNa(numIn)
	select case numIn
		case -1
			yesnona = "YES"
		case 0
			yesnona = "NO"	
		case else
			yesnona = "N/A"
	end select
end function

function YesNoOnly(numIn)
	select case numIn
		case -1
			yesnona = "YES"
		case 0
			yesnona = "NO"	
		case else
			yesnona = "NO"
	end select
end function


%>