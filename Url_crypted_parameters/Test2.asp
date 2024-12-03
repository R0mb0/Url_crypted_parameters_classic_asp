<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="url_crypted_parameters.class.asp"-->
<%
Response.Write("--- Starting second test --- <br><br>")

    Response.Write("--- Initialize and test dictionary status --- <br><br>")

    Dim url 
    Set url = new url_crypted

    Response.Write("--- Set password --- <br><br>")

    url.set_password("Banana")

    Response.Write("--- Get actual URL --- <br><br>")

    Dim actual_url
    actual_url =  url.get_current_url()

    Response.write("URL: " & actual_url & "<br>")

    Response.Write("--- Decrypt actual url params--- <br><br>")
    url.decrypt_actual_params()

    Response.Write("Write parameters <br>")

    url.write_parameters()

    Response.Write("--- Decrypt url params--- <br><br>")
    url.decrypt_url_params(actual_url)

    Response.Write("Write parameters <br>")

    url.write_parameters()
%>