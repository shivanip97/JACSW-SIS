<%
Session("UserLoggedIn") = "false"
Session("AccessLevel") = "3"
Session("Username") = " "
Response.Redirect "index.asp"
%>