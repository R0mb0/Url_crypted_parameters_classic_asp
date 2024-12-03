<%@LANGUAGE="VBSCRIPT"%>
<!--#include file="url_crypted_parameters.class.asp"-->
<%
Response.Write("<h3>--- Starting second test --- </h3><br><br>")

    Response.Write("<h3>--- Initialize and test dictionary status --- </h3><br><br>")

    Dim url 
    Set url = new url_crypted

    Response.Write("<h3>--- Set password --- </h3><br>")

    url.set_password("Banana")

    Response.Write("Password setted: " & url.get_password & "<br><br>")

    Response.Write("<h3>--- Get actual URL --- </h3><br>")

    Dim actual_url
    actual_url =  url.get_current_url()

    Response.write("URL: " & actual_url & "<br><br>")

    Response.Write("<h3>--- Decrypt actual url params --- </h3><br>")
    url.decrypt_actual_params()

    Response.Write("Write parameters <br>")

    url.write_parameters()

    Response.Write("<br><h3>--- Decrypt url params --- </h3><br>")
    url.decrypt_url_params(actual_url)

    Response.Write("Write parameters <br>")

    url.write_parameters()

    Response.Write("<br><h3>--- Get id argument --- </h3><br>")

    Response.Write("Id value: " & url.get_parameter_value("id") & "<br>")

    Response.Write("<h3><br> --- Remove password from params --- </h3><br>")
    url.remove_paramater_by_id("password")
    Response.Write("Write parameters <br>")
     url.write_parameters()
%>