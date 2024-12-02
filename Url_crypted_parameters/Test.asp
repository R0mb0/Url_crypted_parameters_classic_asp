<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="url_crypted_parameters.class.asp"-->
<%
    Dim url 
    Set url = new url_crypted

    Response.write("<h1>" & url.get_current_url() & "</h1>")


%> 