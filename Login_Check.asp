<%
If not Session("UserLoggedIn") = "true" then
    Response.Redirect "index.asp?ErrMsg='Your session has timed out. Please login again to continue'"
End If
%>