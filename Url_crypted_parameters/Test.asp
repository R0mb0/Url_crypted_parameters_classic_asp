<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="url_crypted_parameters.class.asp"-->
<%
    Response.Write("--- Starting test --- <br> <br>")

    Response.Write("--- Initialize and test dictionary status --- <br><br>")

    Dim url 
    Set url = new url_crypted

    Response.Write("--- Set password --- <br><br>")

    url.set_password("Banana")

    Response.Write("Password setted: " & url.get_password)

    Response.Write("--- Add parameters --- <br><br>")

    url.add_paramater "id", 1328
    url.add_paramater "password","blablabla" 

    Response.Write("Write parameters <br>")

    url.write_parameters()

    Response.Write("--- Get actual URL --- <br><br>")

    Dim actual_url
    actual_url =  url.get_current_url()

    Response.write("URL: " & actual_url & "<br>")

    Response.Write("--- Generate new url with parameters crypted --- <br><br>")

    Dim link
    link = url.set_parameters_to_url("test2.asp")

    Response.write("URL: " & link & "<br>")

    Response.write("<h3>Now redirecting!</h3><br>")
    url.redirect(link)

%> 