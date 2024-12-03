<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="url_crypted_parameters.class.asp"-->
<%
    Response.Write("<h3>--- Starting test --- </h3><br><br>")

    Response.Write("<h3>--- Initialize and test dictionary status --- </h3><br><br>")

    Dim url 
    Set url = new url_crypted

    Response.Write("<h3>--- Set password --- </h3><br>")

    url.set_password("Banana")

    Response.Write("Password setted: " & url.get_password & "<br><br>")

    Response.Write("<h3>--- Add parameters --- </h3><br>")

    url.add_paramater "id", 1328
    url.add_paramater "password","blablabla" 

    Response.Write("Write parameters <br>")

    url.write_parameters()

    Response.Write("<br><h3>--- Get actual URL --- </h3><br>")

    Dim actual_url
    actual_url =  url.get_current_url()

    Response.write("URL: " & actual_url & "<br><br>")

    Response.Write("<h3>--- Generate new url with parameters crypted --- </h3><br>")

    Dim link
    link = url.set_parameters_to_url("test2.asp")

    Response.write("URL: " & link & "<br><br>")

    Response.write("<h3>Now redirecting!</h3><br>")
    url.redirect(link)

%> 