<!--#include file="dictionary.class.asp"-->
<!--#include file="rijndael.asp"-->
<% 
Class url_crypted

    Dim my_password
    Dim my_dictionary

    ' Initialization and destruction'
	sub class_initialize()
        my_password = Null 
        Set my_dictionary = new dictionary
	end sub
	
	sub class_terminate()
		my_password = Null 
	end sub

    'Function to add parameters to pass from url '
    Public Function add_paramater(ByVal id, ByVal value)
        my_dictionary.add_element id, value
    End Function 

    'Function to remove parameters by id from paramters to pass from url '
    Public Function remove_paramater_by_id(ByVal id)
        my_dictionary.remove_element_from_key (id)
    End Function 

    'Function to retrieve parameter value from id
    Public Function get_parameter_value(ByVal id)
        get_parameter_value = my_dictionary.get_value_from_key(id)
    End Function 

    Public Function write_parameters()
        my_dictionary.write()
    End Function 

    'Function to set passord to crypt 
    Public Function set_password(ByVal password)
        my_password = password
    End Function 

    'Function to get passord to use to crypt
    Public Function get_password()
        get_password = my_password
    End Function 

    Public Function is_password_setted()
        If IsNull(my_password) Then 
            is_password_setted = False 
        Else 
            is_password_setted = True 
        End if 
    End Function 

    'Function to get current url'
    Public Function get_current_url()
        Dim protocol
        Dim domainName
        Dim fileName
        Dim queryString
        Dim url

        protocol = "http" 
        If lcase(request.ServerVariables("HTTPS"))<> "off" Then 
            protocol = "https" 
        End If

        domainName = Request.ServerVariables("SERVER_NAME") 
        fileName = Request.ServerVariables("SCRIPT_NAME") 
        queryString = Request.ServerVariables("QUERY_STRING")

        url = protocol & "://" & domainName & fileName
        If Len(queryString)<>0 Then
            url = url & "?" & queryString
        End If

        get_current_url = url 
    End Function

    Public Function set_parameters_to_url(url)
        If my_dictionary.get_dimension() > 0 Then 
            If is_password_setted() Then  
                Dim temp 
                Dim is_first
                is_first = true
                Dim my_url 
                my_url = url
                my_url = my_url + "?"
                For Each temp in my_dictionary.get_dictionary()
                    If is_first Then 
                        my_url = my_url + EncryptData temp(0), my_password + "=" + EncryptData temp(1), my_password
                        is_first = false
                    Else 
                        my_url = my_url + "&" + EncryptData temp(0), my_password + "=" + EncryptData temp(1), my_password
                    End If 
                Next 
                set_parameters_to_url = my_url 
            Else 
                Call Err.Raise(vbObjectError + 10, "url_crypted_parameters.class - set_parameters_to_url", "The password is not setted")
            End If 
        Else
            Call Err.Raise(vbObjectError + 10, "url_crypted_parameters.class - set_parameters_to_url", "No parameters to set")
        End If 
    End Function 

End Class 
%>