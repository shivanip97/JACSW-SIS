<fieldset class="step">
       <table id="studentsTable">
	<thead>
		<tr>
            <th align="center">UIN</th>
            <th align="center">First Name</th>
            <th align="center">Last Name</th>
            <th align="center">Maiden Name</th>
            <th align="center">Email</th>
         </tr>
	</thead>
    <tbody>
     <%
					set rs=Server.CreateObject("ADODB.recordset")
					course_query="select * from Applicants order by LastName"
					rs.Open course_query,conn 
                    If rs.EOF Then
                    Response.write("No Students Found")
                    Else
                   Do While NOT rs.Eof  
                    uin = rs("UIN")
           %>
		<tr>
            <td align="center"><div> <% Response.write rs("UIN") %></div></td>
			<td align="center"><div class="edit" id="<%= uin%> FirstName"><% Response.write rs("FirstName") %></div></td>
            <td align="center"><div class="edit" id="<%= uin%> LastName"> <% Response.write rs("LastName") %></div></td>
            <td align="center"><div class="edit" id="<%= uin%> MaidenName"> <% Response.write rs("MaidenName") %></div></td>
            <td align="center"><div class="edit" id="<%= uin%> email"> <% Response.write rs("email") %></div></td>
            <td><button type="submit" name="Button1" onclick="this.form.action='ShowStudentRecords.asp'; this.forms.submit();" id="Button1" value='<% Response.write rs("UIN") %>'>View Records</button></td>
         </tr>
		 <% rs.MoveNext   
        Loop End If %>
	</tbody>
    <% rs.Close
                Set rs=Nothing
                %>
</table>
    </fieldset>