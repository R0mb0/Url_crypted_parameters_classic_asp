# Url crypted parameters in Classic asp

[![Codacy Badge](https://app.codacy.com/project/badge/Grade/9e53d47ef5cc45c8ab5f3305e7918ca3)](https://app.codacy.com/gh/R0mb0/Url_crypted_parameters_classic_asp/dashboard?utm_source=gh&utm_medium=referral&utm_content=&utm_campaign=Badge_grade)

[![Maintenance](https://img.shields.io/badge/Maintained%3F-yes-green.svg)](https://github.com/R0mb0/Url_crypted_parameters_classic_asp)
[![Open Source Love svg3](https://badges.frapsoft.com/os/v3/open-source.svg?v=103)](https://github.com/R0mb0/Url_crypted_parameters_classic_asp)
[![MIT](https://img.shields.io/badge/License-MIT-blue.svg)](https://opensource.org/license/mit)

[![Donate](https://img.shields.io/badge/PayPal-Donate%20to%20Author-blue.svg)](http://paypal.me/R0mb0)

## ⚠️ Dependencies ⚠️

This class needs: `rijndael.asp` and `dictionary.class.asp` to work correctly, the two files must be in the same dir of the class. 

## `url_crypted_parameters.class.asp`'s avaible functions

  - Set the password to use for crypting -> `Public Function set_password(ByVal password)`
  - Get the setted password -> `Public Function get_password()`
  - Check if password is setted -> `Public Function is_password_setted()`
  - Add parameter to crypt -> `Public Function add_paramater(ByVal id, ByVal value)`
  - Change parameter value from id -> `Public Function change_parameter(ByVal id, ByVal value)`
  - Remove parameters -> `Public Function remove_paramater_by_id(ByVal id)`
  - Get parameter value from id -> `Public Function get_parameter_value(ByVal id)`
  - Write all parameters inserted -> `Public Function write_parameters()`
  - Get actual page URL -> `Public Function get_current_url()`
  - Add crypted parameters in URl -> `Public Function set_parameters_to_url(ByVal url)`
  - Redirect a URL to new tab -> `Public Function redirect(ByVal url)`
  - Decrypt actual URL parameters -> `Public Function decrypt_actual_params()`
  - Decrypt URL parameters -> `Public Function decrypt_url_params(ByVal url)`

## How to use

### Page where crypt params 

> From `Test.asp`

1. Initialize the class
   ```
   <%@LANGUAGE="VBSCRIPT"%>
   <!--#include file="url_crypted_parameters.class.asp"-->
   <%
    Dim url 
    Set url = new url_crypted

    url.set_password("Banana")
   ```
2. Add params to crypt in URL
   ```
    url.add_paramater "id", 1328
    url.add_paramater "password","blablabla"
   ```
3. Generate URL with crypted params
   ```
    Dim link
    link = url.set_parameters_to_url("test2.asp")
   %>
   ```

### Page where decrypt params 

> From `Test2.asp`

1. Initialize the class
   ```
   <%@LANGUAGE="VBSCRIPT"%>
   <!--#include file="url_crypted_parameters.class.asp"-->
   <%
    Dim url 
    Set url = new url_crypted

    url.set_password("Banana")
   ```
2. Decrypt page params
   ```
    url.decrypt_actual_params()
   ```
3. Access information
   ```
    Response.Write("Id value: " & url.get_parameter_value("id") & "<br>")
   <%
   ```
