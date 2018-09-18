<%@LANGUAGE="VBSCRIPT" CODEPAGE="65001" enablesessionstate="True" %>
<!--#include file="Login_Check.asp"-->

<!--#include file="DBconn.asp"-->
<% 

LastUpdatedTime = Time()
LastUpdatedDt = date()    
sub Write_CSV_From_Recordset(RS)

    '
    ' Export Recordset to CSV
    ' http://salman-w.blogspot.com/2009/07/export-recordset-data-to-csv-using.html
    '
    ' This sub-routine Response.Writes the content of an ADODB.RECORDSET in CSV format
    ' The function closely follows the recommendations described in RFC 4180:
    ' Common Format and MIME Type for Comma-Separated Values (CSV) Files
    ' http://tools.ietf.org/html/rfc4180
    '
    ' @RS: A reference to an open ADODB.RECORDSET object
    '

    if RS.EOF then
    
        '
        ' There is no data to be written
        '
        exit sub
    
    end if

    dim RX
    set RX = new RegExp
        RX.Pattern = "\r|\n|,|"""

    dim i
    dim Field
    dim Separator

    '
    ' Writing the header row (header row contains field names)
    '

    Separator = ""
    for i = 0 to RS.Fields.Count - 1
        Field = RS.Fields(i).Name
        if RX.Test(Field) then
            '
            ' According to recommendations:
            ' - Fields that contain CR/LF, Comma or Double-quote should be enclosed in double-quotes
            ' - Double-quote itself must be escaped by preceeding with another double-quote
            '
            Field = """" & Replace(Field, """", """""") & """"
        end if
        Response.Write Separator & Field
        Separator = ","
    next
    Response.Write vbNewLine

    '
    ' Writing the data rows
    '

    do until RS.EOF
        Separator = ""
        for i = 0 to RS.Fields.Count - 1
            '
            ' Note the concatenation with empty string below
            ' This assures that NULL values are converted to empty string
            '
            Field = RS.Fields(i).Value & ""
            if RX.Test(Field) then
                Field = """" & Replace(Field, """", """""") & """"
            end if
            Response.Write Separator & Field
            Separator = ","
        next
        Response.Write vbNewLine
        RS.MoveNext
    loop

end sub

'
' EXAMPLE USAGE
'
' - Open a RECORDSET object (forward-only, read-only recommended)
' - Send appropriate response headers
' - Call the function
'

dim RS1
set RS1 = Server.CreateObject("ADODB.recordset")
    RS1.Open "SELECT distinct c.UIN,  c.DegreeProgram, c.ProgramType, c.Track, c.Decision,  isnull(a.Application_Status,'') As Application_Status, max(isnull(a.OAR_Application_Date ,'')) As Date_of_Initial_Entry, isnull(a.ReadyforreviewDate,'') As Date_Received, isnull(c.TermGraduated,'') as TermGraduated, c.Confirmed, c.AdmitTerm, c.Race_ethinicity, c.Gender, c.DateOfBirth, c.Concentration, c.EMail, c.Personalemail, c.LastName, c.FirstName, isnull(c.Salutation,'') as Salutation, c.CurrentAddress1, c.CurrentAddress2,c.CurrentCity, c.CurrentState, c.CurrentZipCode from CurrentStudents c left outer join Applicants a on c.UIN = a.UIN and a.Application_Status is not null and a.Application_Status != 'IN-Incomplete' and c.ProgramType <> '' and c.AdmitTerm <> '' and (c.Graduated != 'Y'  or c.Graduated is null) and c.Status not in ('LOA', 'TRANS','WDN') and c.Decision <> 'DF' Group By c.UIN,  c.DegreeProgram, c.ProgramType,c.Decision,c.EMail,c.Track,c.Personalemail,  a.Application_Status, a.ReadyforreviewDate , c.TermGraduated, c.Confirmed, c.AdmitTerm, c.Race_ethinicity, c.Gender, c.DateOfBirth, c.Concentration, c.LastName, c.FirstName, c.Salutation, c.CurrentAddress1, c.CurrentAddress2,c.CurrentCity, c.CurrentState, c.CurrentZipCode order by 14", conn, 0, 1

Response.ContentType = "text/csv"

Response.AddHeader "Content-Disposition", "filename=CurrentStudentsReport.csv"


Write_CSV_From_Recordset RS1
   
%>
