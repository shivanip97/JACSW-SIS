<%
Application("previousyear")="2012"
Application("lastyear")="2013"
Application("currentyear")="2014"
Application("nextyear")="2015"
set conn=server.CreateObject("ADODB.connection")
conn.ConnectionString="Provider=SQLOLEDB;Data Source=cc-sql2k14-jasw\JACSW,59048;Initial Catalog=JACSWStudent;User Id=JACSW-webapp;Password=JACSWw3b@pp!;"
conn.Open
set conn1=server.CreateObject("ADODB.connection")
conn1.ConnectionString="Provider=SQLOLEDB;Data Source=cc-sql2k14-jasw\JACSW,59048;Initial Catalog=JACSWField;User Id=JACSW-webapp;Password=JACSWw3b@pp!;"
conn1.Open
%>