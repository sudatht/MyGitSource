<%@ CODEPAGE=65001 %> 

<% Response.Charset="UTF-8" %> 
<%
   dim CESecurity, CESecurityArray
   dim MaxImageSize
   dim MaxMediaSize
   dim MaxFlashSize
   dim MaxDocumentSize
   dim MaxTemplateSize
   dim ImageGalleryPath
   dim MediaGalleryPath
   dim FlashGalleryPath
   dim TemplateGalleryPath
   dim FilesGalleryPath
   dim AllowUpload
   dim AllowCreateFolder
   dim AllowRename
   dim AllowDelete
   dim ImageFilters
   dim MediaFilters
   dim DocumentFilters
   dim TemplateFilters
   dim Culture
   dim nl
   dim Base64EncMap(63)
   dim Base64DecMap(127)
   call initCodecs
  
   dim badrequest
   
   s_CESecurity=Trim(Session("CESecurity"))   
   c_CESecurity=Trim(Request.Cookies("CESecurity"))
   q_CESecurity=Trim(Request.QueryString("setting"))
   
   badrequest=false
   
   if s_CESecurity <> "" Then  
       if s_CESecurity <> q_CESecurity Then 
            badrequest = true
       End if   
   else   
       if c_CESecurity <> q_CESecurity Then 
            badrequest = true
       End if
   end if
   
   If badrequest then   
       Response.Write "The area you are attempting to access is forbidden"	
	   Response.End        
   end if
   
   CESecurity = d(q_CESecurity)
   if CESecurity <> "" Then
        CESecurityArray = Split(CESecurity,"|")
        Culture=CESecurityArray(18)
   Else
        Culture="en-US"      
   End if 
   
   dim Theme
   Theme = Trim(Request.QueryString("Theme"))
   
    Public Function GetString(instring)
	    dim t
    	
	    t = GetStringByCulture(instring,Culture)
    	
	    If t = ""  then
		    t= GetStringByCulture(instring,"_default")
	    End If
    	
	    If t = ""  then
		    t= "{"&instring&"}"	
	    End If
	    GetString= t	
    End Function
    
    
    Dim path
    path=Request.ServerVariables("SCRIPT_NAME")
    path=left(path,InStrRev(path,"/")-8)   
  
    Public Function GetStringByCulture(instring,input_culture)
	    dim scriptname,xmlfilename,doc,temp
	    dim node,selectednode,optionnodelist,errobj
	    dim selectednodes

	    xmlfilename= Server.MapPath(path&"languages/"&input_culture&".xml")

	    ' Create an object to hold the XML
	    set doc = server.CreateObject("Microsoft.XMLDOM")

	    ' For ASP, wait until the XML is all ready before continuing
	    doc.async = False

	    ' Load the XML file or return an error message and stop the script
	    if not Doc.Load(xmlfilename) then
		    Response.Write "Failed to load the language text from the XML file.<br>"
		    Response.End
	    end if

	    ' Make sure that the interpreter knows that we are using XPath as our selection language
	    doc.setProperty "SelectionLanguage", "XPath"

	    set selectednode= doc.selectSingleNode("/resources/resource[@name='"&instring&"']")
	    if IsObject(selectednode) and not selectednode is nothing  then
		    GetStringByCulture=CuteEditor_Decode(selectednode.text)
	    else
		    GetStringByCulture=""		
	    end if
    End Function    
    PUBLIC FUNCTION d(scrambled)
          if len(scrambled) = 0 then
               d = ""
               exit function
          end if
          ' ignore padding
          dim realLen
          realLen = len(scrambled)
          do while mid(scrambled, realLen, 1) = "="
               realLen = realLen - 1
          loop
          dim ret, ndx, by4, first, second, third, fourth
          ret = ""
          by4 = (realLen \ 4) * 4
          ndx = 1
          do while ndx <= by4
               first  = Base64DecMap(asc(mid(scrambled, ndx+0, 1)))
               second = Base64DecMap(asc(mid(scrambled, ndx+1, 1)))
               third  = Base64DecMap(asc(mid(scrambled, ndx+2, 1)))
               fourth = Base64DecMap(asc(mid(scrambled, ndx+3, 1)))
               ret = ret & chr( ((first * 4) AND 255) +   ((second \ 16) AND 3))
               ret = ret & chr( ((second * 16) AND 255) + ((third \ 4) AND 15) )
               ret = ret & chr( ((third * 64) AND 255) +  (fourth AND 63) )
               ndx = ndx + 4
          loop
          if ndx < realLen then
               first  = Base64DecMap(asc(mid(scrambled, ndx+0, 1)))
               second = Base64DecMap(asc(mid(scrambled, ndx+1, 1)))
               ret = ret & chr( ((first * 4) AND 255) +   ((second \ 16) AND 3))
               if realLen MOD 4 = 3 then
                    third = Base64DecMap(asc(mid(scrambled,ndx+2,1)))
                    ret = ret & chr( ((second * 16) AND 255) + ((third \ 4) AND 15) )
               end if
          end if

          d = ret
     END FUNCTION
     PUBLIC SUB initCodecs()
          ' init vars
          nl = "<P>" & chr(13) & chr(10)
          ' setup base 64
          dim max, idx
          const BASE_64_MAP_INIT ="ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
          max = len(BASE_64_MAP_INIT)
          for idx = 0 to max - 1
               ' one based string
               Base64EncMap(idx) = mid(BASE_64_MAP_INIT, idx + 1, 1)
          next
          for idx = 0 to max - 1
               Base64DecMap(ASC(Base64EncMap(idx))) = idx
          next
     END SUB
     
     PUBLIC FUNCTION CuteEditor_Decode(input_str)        
	    input_str=Replace(input_str,"#1","<")
	    input_str=Replace(input_str,"#2",">")
	    input_str=Replace(input_str,"#3","&")
	    input_str=Replace(input_str,"#4","*")
	    input_str=Replace(input_str,"#5","o")
	    input_str=Replace(input_str,"#6","O")
	    input_str=Replace(input_str,"#7","s")
	    input_str=Replace(input_str,"#8","S")
	    input_str=Replace(input_str,"#9","e")
	    input_str=Replace(input_str,"#a","E")
	    input_str=Replace(input_str,"#0","#")
	    CuteEditor_Decode = input_str
     END FUNCTION
   dim setting
   setting = Trim(Request.QueryString("setting")) 
   setting="setting="+setting
%>