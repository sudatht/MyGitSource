<%
Class CuteEditor
    private d    
	private s_AccessKey
	private s_BackColor
	private s_BorderStyle
	private s_BorderWidth
	private s_BorderColor
	private s_RenderRichDropDown
	private s_ContextMenuMode
	private s_URLType
	private s_EmptyAlternateText
	private s_HyperlinkTarget
	private s_ServerName
	private s_BrowserType
	private s_activetab
	private s_AllowEditServerSideCode
	private s_AllowPasteHtml
	private s_autoconfigure
	private s_AutoParseClasses
	private s_BaseHref
	private s_breakelement
	private s_ResizeMode
	private s_CodeViewTemplateItemList
	private s_ConfigurationPath
	private s_ConvertHTMLTagstoLowercase
	private s_CustomCulture
	private s_DisableAutoFormatting
	private s_DisableClassList 
	private s_disableitemlist
	private s_DOCTYPE
	private s_DownLevelColumns
	private s_TabSpaces
	private s_ResizeStep
	private s_DownLevelRows 
	private s_EditCompleteDocument
	private s_EditorBodyStyle
	private s_EditorBodyId
	private s_EditorBodyClass
	private s_EditorOnPaste
	private s_EditorWysiwygModeCss
	private s_PreviewModeCss
	private s_EnableAntiSpamEmailEncoder
	private s_EnableBrowserContextMenu
	private s_enableclientscript
	private s_EnableContextMenu
	private	s_EnableStripScriptTags
	private s_filespath	
	private s_Focus
	private s_FullPage
	private s_ToggleBorder
	private s_height
	private s_HelpUrl
	Public ID
	private s_MaxHTMLLength
	private s_MaxTextLength
	private s_PrintFullWebPage
	private s_ReadOnly
	private s_removeservernamesfromurl 
	private s_RemoveTBODYTag
	private s_SecurityPolicyFile
	private s_showBottomBar
	private s_ShowCodeViewToolBar
	private s_ShowDecreaseButton
	private s_ShowEnlargeButton
	private s_showgroupmenuimage
	private s_showhtmlmode
	private s_ShowTagSelector
	private s_ShowWordCount
	private s_showpreviewmode
	private s_ShowToolBar
	private s_UseStandardDialog
	private s_tabindex
	private s_templateitemlist
	private s_TextAreaStyle
	private s_ThemeType
	private s_UseFontTags
	private s_UseHTMLEntities
	private s_UsePhysicalFormattingTags
	private s_UseRelativeLinks
	private s_UseSimpleAmpersand
	private s_width	
	private s_cssclassdropdownMenuNames
	private s_cssclassdropdownMenuList
	private s_inlinestyledropdownMenuNames
	private s_inlinestyledropdownMenuList
	private s_ParagraphsListMenuNames
	private s_ParagraphsListMenuList
	private s_FontFacesList
	private s_FontSizesList
	private s_linksdropdownMenuList
	private s_linksdropdownMenuNames
	private s_codesnippetdropdownMenuNames
	private s_codesnippetdropdownMenuList
	private s_imagesdropdownMenuNames
	private s_imagesdropdownMenuList
	private s_maxImageSize
	private s_maxMediaSize
	private s_maxFlashSize
	private s_maxDocumentSize
	private s_MaxTemplateSize
	private s_AllowUpload
	private s_AllowCreateFolder
	private s_AllowRename
	private s_AllowDelete
	private s_imagegallerypath	
	private s_MediaGalleryPath
	private s_FlashGalleryPath
	private s_TemplateGalleryPath
	private s_FilesGalleryPath
	private s_ImageFilters
	private s_MediaFilters
    private s_DocumentFilters
    private s_TemplateFilters
    private s_XMLOutput
	
	private s_Text
	private s_flashpath	
	private s_mediapath	
	private s_documentpath	
	private s_subsequent
	private s_ParagraphsDropDownWidth
	private s_SizesDropDownWidth	
	private s_ZoomsDropDownWidth	
	private s_StylesDropDownWidth
	private s_CodeSnippetsDropDownWidth
	private s_ImagesDropDownWidth
	private s_FontsDropDownWidth
	private s_LinksDropDownWidth
	private s_ZoomsList
	
	private s_MaintainAspectRatioWhenDraggingImage
	private s_EnableObjectResizing
	private s_EnableDragDrop
	
	Public Property Let MaintainAspectRatioWhenDraggingImage(yesornot)
		s_MaintainAspectRatioWhenDraggingImage = Lcase(CStr(yesornot))
	End Property	
	Public Property Let EnableObjectResizing(yesornot)	
		s_EnableObjectResizing =  Lcase(CStr(yesornot))	
	End Property
	Public Property Let EnableDragDrop(yesornot)	
		s_EnableDragDrop =  Lcase(CStr(yesornot))	
	End Property	
	Public Property Let AccessKey(s)
		s_AccessKey = s
	End Property	
	
    'version 5.0
	Public Property Let ActiveTab(input_string)	
		s_activetab =  input_string	
	End Property
	Public Property Let AllowEditServerSideCode (yesornot)
		s_AllowEditServerSideCode = Lcase(CStr(yesornot))
	End Property
	Public Property Let AllowPasteHtml(yesornot)
		s_AllowPasteHtml = Lcase(CStr(yesornot))
	End Property	
	Public Property Let AutoConfigure(cfg)
		s_autoconfigure = cfg
	End Property
	Public Property Let AutoParseClasses (yesornot)
		s_AutoParseClasses = Lcase(CStr(yesornot))
	End Property
	Public Property Let BaseHref(input_string)
		s_BaseHref = input_string
	End Property
	Public Property Let BreakElement(input_string)
		s_breakelement = input_string
	End Property	
	Public Property Let ResizeMode(input_string)
		s_ResizeMode = input_string
	End Property		
	Public Property Let CodeViewTemplateItemList(Input_list)
		s_CodeViewTemplateItemList = Input_list
	End Property		
	Public Property Let ConfigurationPath(s)
		s_ConfigurationPath = s
	End Property		
	Public Property Let ConvertHTMLTagstoLowercase(yesornot)
		s_ConvertHTMLTagstoLowercase = Lcase(CStr(yesornot))
	End Property
	Public Property Let DisableAutoFormatting(yesornot)
		s_DisableAutoFormatting = Lcase(CStr(yesornot))
	End Property	
	Public Property Let DisableClassList(yesornot)
		s_DisableClassList = Lcase(CStr(yesornot))
	End Property			
	Public Property Let DisableItemList(Input_list)
		s_disableitemlist = Input_list
	End Property		
	Public Property Let DOCTYPE(s)
		s_DOCTYPE = s
	End Property	
	Public Property Let BackColor(s)
		s_BackColor = s
	End Property	
	Public Property Let BorderColor (s)
		s_BorderColor = s
	End Property
	Public Property Let BorderStyle(s)
		s_BorderStyle = s
	End Property
	Public Property Let BorderWidth(s)
		s_BorderWidth = Cint(s)
	End Property
	Public Property Let DownLevelColumns(n)
		s_DownLevelColumns = n
	End Property	
	Public Property Let TabSpaces(n)
		s_TabSpaces = n
	End Property	
	Public Property Let ResizeStep(n)
		s_ResizeStep = n
	End Property
	Public Property Let DownLevelRows (n)
		s_DownLevelRows  = n
	End Property	
	Public Property Let EditCompleteDocument (yesornot)
		s_EditCompleteDocument = Lcase(CStr(yesornot))
	End Property			
	Public Property Let EditorBodyStyle(input_string)
		s_EditorBodyStyle = input_string
	End Property		
	Public Property Let EditorBodyId(input_string)
		s_EditorBodyId = input_string
	End Property		
	Public Property Let EditorBodyClass(input_string)
		s_EditorBodyClass = input_string
	End Property
	Public Property Let ContextMenuMode(input_string)
		s_ContextMenuMode= CStr(input_string)
	End Property	
	Public Property Let EditorOnPaste(input_string)
		s_EditorOnPaste= CStr(input_string)
	End Property	
	Public Property Let EmptyAlternateText(input_string)
		s_EmptyAlternateText= CStr(input_string)
	End Property	
	Public Property Let HyperlinkTarget(input_string)
		s_HyperlinkTarget= CStr(input_string)
	End Property	
	Public Property Let URLType(input_string)
		s_URLType= CStr(input_string)
	End Property
	Public Property Let EditorWysiwygModeCss(input_string)
		s_EditorWysiwygModeCss = input_string
	End Property
	Public Property Let PreviewModeCss(input_string)
		s_PreviewModeCss = input_string
	End Property
	Public Property Let EnableAntiSpamEmailEncoder(yesornot)
		s_EnableAntiSpamEmailEncoder = Lcase(CStr(yesornot))
	End Property	
	Public Property Let EnableBrowserContextMenu (yesornot)
		s_EnableBrowserContextMenu = Lcase(CStr(yesornot))
	End Property		
	Public Property Let EnableClientScript(trueornot)
		s_enableclientscript = trueornot
	End Property		
	Public Property Let EnableContextMenu(trueornot)
		s_EnableContextMenu = trueornot
	End Property
	Public Property Let EnableStripScriptTags(yesornot)
		s_EnableStripScriptTags = Lcase(CStr(yesornot))
	End Property			
	Public Property Let FilesPath(s)
		if Right(s,1) = "/" then
			s = Left(s, Len(s)-1)
		end if
		s_filespath = s
	End Property	
	Public Property Let Focus(yesornot)
		s_Focus= Lcase(CStr(yesornot))
	End Property
	Public Property Let FullPage(yesornot)
		s_FullPage= Lcase(CStr(yesornot))
	End Property
	Public Property Let ToggleBorder(yesornot)
		s_ToggleBorder= Lcase(CStr(yesornot))
	End Property
	Public Property Let Height(n)
		s_height = n
	End Property	
	Public Property Let HelpUrl(s)
		s_HelpUrl = s
	End Property	
	Public Property Let MaxHTMLLength(n)
		s_MaxHTMLLength = n
	End Property		
	Public Property Let MaxTextLength(n)
		s_MaxTextLength = n
	End Property		
	Public Property Let PrintFullWebPage(trueornot)
		s_PrintFullWebPage = trueornot
	End Property
	Public Property Let ReadOnly(yesornot)
		s_ReadOnly = Lcase(CStr(yesornot))
	End Property		
	Public Property Let RemoveServerNamesFromUrl(yesornot)
		s_removeservernamesfromurl  = Lcase(CStr(yesornot))
	End Property		
	Public Property Let RemoveTBODYTag(yesornot)
		s_RemoveTBODYTag = Lcase(CStr(yesornot))
	End Property	
	Public Property Let SecurityPolicyFile (s)
		s_SecurityPolicyFile = s
	End Property	
	Public Property Let ShowBottomBar(yesornot)
		s_showBottomBar = Lcase(CStr(yesornot))
	End Property
	Public Property Let ShowCodeViewToolBar(yesornot)
		s_ShowCodeViewToolBar = Lcase(CStr(yesornot))
	End Property	
	Public Property Let ShowDecreaseButton(yesornot)
		s_ShowDecreaseButton = Lcase(CStr(yesornot))
	End Property	
	Public Property Let ShowEnlargeButton(yesornot)
		s_ShowEnlargeButton = Lcase(CStr(yesornot))
	End Property		
	Public Property Let ShowGroupMenuImage(yesornot)
		s_showgroupmenuimage = Lcase(CStr(yesornot))
	End Property
	Public Property Let ShowHtmlMode(yesornot)
		s_showhtmlmode = Lcase(CStr(yesornot))
	End Property	
	Public Property Let ShowTagSelector(yesornot)
		s_ShowTagSelector = Lcase(CStr(yesornot))
	End Property
	Public Property Let ShowWordCount(yesornot)
		s_ShowWordCount = Lcase(CStr(yesornot))
	End Property
	Public Property Let UseStandardDialog(yesornot)
		s_UseStandardDialog = Lcase(CStr(yesornot))
	End Property
	Public Property Let ShowPreviewMode(yesornot)
		s_showpreviewmode = Lcase(CStr(yesornot))
	End Property
	Public Property Let ShowToolBar(yesornot)
		s_ShowToolBar = Lcase(CStr(yesornot))
	End Property
	Public Property Let TabIndex(i)
		s_tabindex = i
	End Property	
	Public Property Let TemplateItemList(s)
		s_templateitemlist = s
	End Property
	Public Property Let TextAreaStyle(s)
		s_TextAreaStyle = s
	End Property
	Public Property Let ThemeType(Input_theme)
		s_ThemeType = Input_theme
	End Property
	Public Property Let UseFontTags (yesornot)
		s_UseFontTags = Lcase(CStr(yesornot))
	End Property
	Public Property Let RenderRichDropdown (yesornot)
	    s_RenderRichDropdown = Lcase(CStr(yesornot))
	End Property
	Public Property Let UseHTMLEntities (yesornot)
		s_UseHTMLEntities = Lcase(CStr(yesornot))
	End Property
	Public Property Let UsePhysicalFormattingTags (yesornot)
		s_UsePhysicalFormattingTags = Lcase(CStr(yesornot))
	End Property	
	Public Property Let UseSimpleAmpersand(yesornot)
		s_UseSimpleAmpersand= Lcase(CStr(yesornot))
	End Property	
	Public Property Let Width(n)
		s_width = n
	End Property					
	Public Property Let CssClassStyleDropDownMenuNames(input_string)
		s_cssclassdropdownMenuNames = input_string
	End Property	
	Public Property Let CssClassStyleDropDownMenuList(input_string)
		s_cssclassdropdownMenuList = input_string
	End Property					
	Public Property Let InlineStyleDropDownMenuNames(input_string)
		s_inlinestyledropdownMenuNames = input_string
	End Property	
	Public Property Let InlineStyleDropDownMenuList(input_string)
		s_inlinestyledropdownMenuList = input_string
	End Property	
	Public Property Let ParagraphsListMenuNames(list)
		s_ParagraphsListMenuNames = list
	End Property	
	Public Property Let ParagraphsListMenuList(list)
		s_ParagraphsListMenuList = list
	End Property			
	Public Property Let FontFacesList(list)
		s_FontFacesList = list
	End Property
	Public Property Let FontSizesList(list)
		s_FontSizesList = list
	End Property	
	Public Property Let LinksDropDownMenuList(list)
		s_linksdropdownMenuList = list
	End Property		
	Public Property Let LinksDropDownMenuNames(list)
		s_linksdropdownMenuNames = list
	End Property		
	Public Property Let CodeSnippetDropDownMenuNames(input_string)
		s_codesnippetdropdownMenuNames = input_string
	End Property	
	Public Property Let CodeSnippetDropDownMenuList(input_string)
		s_codesnippetdropdownMenuList = input_string
	End Property		
	Public Property Let ImagesDropDownMenuNames(input_string)
		s_imagesdropdownMenuNames = input_string
	End Property	
	Public Property Let ImagesDropDownMenuList(input_string)
		s_imagesdropdownMenuList = input_string
	End Property		
	Public Property Let ZoomsList(list)
		s_ZoomsList = list
	End Property
	Public Property Let MaxImageSize(size)
		s_maxImageSize = size
	End Property		
	Public Property Let MaxMediaSize(size)
		s_maxMediaSize = size
	End Property
	Public Property Let MaxFlashSize(size)
		s_maxFlashSize = size
	End Property	
	Public Property Let MaxDocumentSize(size)
		s_maxDocumentSize = size
	End Property
	Public Property Let MaxTemplateSize(size)
		s_MaxTemplateSize = size
	End Property
	Public Property Let ImageGalleryPath(path)
		s_imagegallerypath = path
	End Property	
	Public Property Let MediaGalleryPath(path)
		s_MediaGalleryPath = path
	End Property	
	Public Property Let FlashGalleryPath(path)
		s_FlashGalleryPath = path
	End Property	
	Public Property Let TemplateGalleryPath(path)
		s_TemplateGalleryPath = path
	End Property		
	Public Property Let FilesGalleryPath(path)
		s_FilesGalleryPath = path
	End Property		
	Public Property Let AllowCreateFolder(yesornot)
		s_AllowCreateFolder = Lcase(CStr(yesornot))
	End Property	
	Public Property Let AllowUpload(yesornot)
		s_AllowUpload = Lcase(CStr(yesornot))
	End Property	
	Public Property Let AllowRename(yesornot)
		s_AllowRename = Lcase(CStr(yesornot))
	End Property	
	Public Property Let AllowDelete(yesornot)
		s_AllowDelete = Lcase(CStr(yesornot))
	End Property			
	Public Property Let ImageFilters(s)
		s_ImageFilters = s
	End Property			
	Public Property Let MediaFilters(s)
		s_MediaFilters = s
	End Property			
	Public Property Let DocumentFilters(s)
		s_DocumentFilters = s
	End Property		
	Public Property Let TemplateFilters(s)
		s_TemplateFilters = s
	End Property	
	
	Public Property Get Text
		Text = Request.Form(ID)
	End Property	
		
	Public CustomAddons		
	Public Property Let Text(initialText)
		s_Text = initialText & ""
		s_Text = Server.HTMLEncode( s_Text )
	End Property		
	Public Property Let UseRelativeLinks(yesornot)
		s_UseRelativeLinks  = Lcase(CStr(yesornot))
	End Property			
	Public Property Let subsequent(yesornot)
		s_subsequent = Lcase(CStr(yesornot))
	End Property
	Public Property Let CustomCulture(s)
		s_CustomCulture= s
	End Property	
	Public Property Let XHTMLOutput (yesornot)
		s_XMLOutput = Lcase(CStr(yesornot))
	End Property	
	
	
	
	Public Property Get ClientID
		ClientID = "CE_"&ID&"_ID"
	End Property			
	
	'********************************************************
	' Begin Event Handlers
	'********************************************************

	private Sub Class_Initialize()
		s_activetab = "Edit"
		s_AllowEditServerSideCode = "false"
		s_AllowPasteHtml = "true"
		s_autoconfigure = "default"
		s_AutoParseClasses = "true"
		s_BaseHref=""
		s_breakelement = "Div"
		s_ResizeMode = "ResizeCorner"
		s_CodeViewTemplateItemList="Save,Print,Cut,Copy,Paste,Find,ToFullPage,FromFullPage,SelectAll,SelectNone"
		s_ConvertHTMLTagstoLowercase  = "true"
		s_CustomCulture = "en-en"
		s_DisableAutoFormatting  = "false"
		s_DownLevelColumns = 50
		s_DownLevelRows  = 13
		s_TabSpaces = 3
		s_ResizeStep = 100
		s_DOCTYPE=""
		s_EditCompleteDocument = "false"
		s_EditorBodyStyle=""
		s_EditorBodyClass=""
		s_EditorBodyId=""
		s_EditorOnPaste = "ConfirmWord"
		s_URLType = "Default"
		s_EmptyAlternateText = "ForceAdd"
		s_HyperlinkTarget = "Default"
		's_ContextMenuMode = "Default"
		s_EditorWysiwygModeCss = ""
		s_PreviewModeCss = ""
		s_EnableAntiSpamEmailEncoder = "true"
		s_EnableBrowserContextMenu = "true"
		s_enableclientscript = "true"
		s_EnableContextMenu = "true"
		s_EnableStripScriptTags  = "false"
		s_filespath="cuteeditor_files" 
		s_Focus = "false"
		s_FullPage = "false"
		s_ToggleBorder = "true"
		s_height = "300"
		s_MaxHTMLLength = 0
		s_MaxTextLength = 0
		s_PrintFullWebPage = "false"
		s_ReadOnly = "false"
		s_removeservernamesfromurl  = "true"
		s_RemoveTBODYTag  = "false"
		s_SecurityPolicyFile = "default.config"
		s_showBottomBar = "true"
		s_ShowCodeViewToolBar = "true"
		s_ShowDecreaseButton = "true"
		s_ShowEnlargeButton = "true"
		s_showgroupmenuimage = "true"
		s_showhtmlmode = "true"
		s_ShowTagSelector= "true"
		s_ShowWordCount= "true"
		s_showpreviewmode = "true"
		s_UseStandardDialog = "false"
		s_ShowToolBar = "true"
		s_ThemeType = "Office2007"
		s_UseFontTags= "false"
		s_RenderRichDropdown = "true"
		s_UseHTMLEntities = "true"
		s_UsePhysicalFormattingTags = "false"
		s_UseRelativeLinks  = "true"
		s_MaintainAspectRatioWhenDraggingImage   = "true"
		s_EnableObjectResizing   = "true"
		s_EnableDragDrop   = "true"
		s_width = "780"
		s_BackColor = "#F4F4F3"
		s_BorderColor = "#dddddd"
		s_BorderStyle  = "Solid"
		s_BorderWidth  = 1
		s_HelpUrl = s_filespath&"/Help/default.htm"
		
		s_subsequent = "false"
		s_UseSimpleAmpersand  = "false"
		s_tabindex = 0
		s_XMLOutput = "false"
		s_ServerName=Request.ServerVariables("SERVER_NAME")	
		Call GetAllIndexMap ()
		Call GetBrowserType ()
		
		set xmldocs=server.CreateObject ("scripting.dictionary")
    
	End Sub 
	
	Sub GetBrowserType ()	
		Dim userAgent
		s_BrowserType = "false"
		userAgent = Request.ServerVariables("HTTP_USER_AGENT")
		if InStr(1, userAgent, "MSIE", 1) > 0 AND InStr(1, userAgent, "Win", 1) > 0 AND InStr(1, userAgent, "Opera", 1) = 0 then
			if Trim(Mid(userAgent, inStr(1, userAgent, "MSIE", 1)+5, 3)) >= "5.5" OR Trim(Mid(userAgent, inStr(1, userAgent, "MSIE", 1)+5, 3)) = "5,5" then
				s_BrowserType = "winIE"
			end if
		ElseIf inStr(1, userAgent, "Firebird", 1) then
			if CLng(Trim(Mid(userAgent, CInt(inStr(1, userAgent, "Gecko/", 1)+6), 8))) >= 20030728 then
					s_BrowserType = "Gecko"
			end if
		ElseIf inStr(1, userAgent, "Gecko", 1) > 0 AND inStr(1, userAgent, "Firebird", 1) = 0 AND isNumeric(Trim(Mid(userAgent, CInt(inStr(1, userAgent, "Gecko/", 1)+6), 8))) then
			if CLng(Trim(Mid(userAgent, CInt(inStr(1, userAgent, "Gecko/", 1)+6), 8))) => 20030312 then
				s_BrowserType = "Gecko"
			end if
		ElseIf inStr(1, userAgent, "Opera", 1) > 0 then
			dim OperaNumber
			OperaNumber = Mid(userAgent, inStr(1, userAgent, "Opera", 1)+6, 1)
			if IsNumeric (OperaNumber) then
				if CInt(OperaNumber) => 9  then
					s_BrowserType = "Opera"
				end if
			end if		    
		ElseIf inStr(1, userAgent, "Safari", 1) > 0  then
			if CInt(Trim(Mid(userAgent, CInt(inStr(1, userAgent, "AppleWebKit/", 1)+12), 3))) => 522 then
				s_BrowserType = "Safari"
			ElseIf CInt(Trim(Mid(userAgent, CInt(inStr(1, userAgent, "AppleWebKit/", 1)+12), 3))) => 312 then
				s_BrowserType = "Safari12"	
			end if
		end if
		If inStr(1, userAgent, "Chrome", 1) > 0 AND inStr(1, userAgent, "Safari", 1) > 0  then
			s_BrowserType = "Gecko"
		end if
		If inStr(1, userAgent, "iphone", 1) > 0 or inStr(1, userAgent, "windows ce", 1) > 0 or inStr(1, userAgent, "blackberry", 1) > 0 or inStr(1, userAgent, "opera mini", 1) > 0 or inStr(1, userAgent, "mobile", 1) > 0 or inStr(1, userAgent, "palm", 1) > 0 or inStr(1, userAgent, "portable", 1) > 0 then
			s_BrowserType = "false"
		End If
	End Sub
	Public Property Get BrowserType
	    BrowserType=s_BrowserType
	End Property
		
	Public Property Get RenderRichDropdown
	    if Lcase(BrowserType) = "safari" or Lcase(BrowserType) = "safari12" then	    
	        RenderRichDropdown="false"
	    else
	        RenderRichDropdown=s_RenderRichDropdown
	    end if
	End Property
	
	private Function GetFormatBlockCode (s_block)
		Select case Lcase(s_block)
			Case "normal":
				GetFormatBlockCode = "<P>"
			Case "heading 1":
				GetFormatBlockCode = "<H1>"
			Case "heading 2":
				GetFormatBlockCode = "<H2>"
			Case "heading 3":
				GetFormatBlockCode = "<H3>"
			Case "heading 4":
				GetFormatBlockCode = "<H4>"
			Case "heading 5":
				GetFormatBlockCode = "<H5>"
			Case "heading 6":
				GetFormatBlockCode = "<H6>"
			Case "address":
				GetFormatBlockCode = "Address"
			Case "formatted":
				GetFormatBlockCode = "Formatted"
			Case "definition term":
				GetFormatBlockCode = "Definition Term"
			Case Else
				GetFormatBlockCode = "<P>"			
		End Select
	End Function

	private Function CreateToolBar()
	    Dim strTemp	
		if s_templateitemlist <> "" then
			strTemp = GetToolbarFromItemList(s_templateitemlist,s_disableitemlist)
		elseif s_ConfigurationPath <> "" then
			strTemp = GetToolbarItems(s_ConfigurationPath,s_disableitemlist)
		else
			strTemp = GetToolbarItems(GetURL("Configuration/AutoConfigure/"&s_autoconfigure&".config"),s_disableitemlist)			
		end if	
		CreateToolBar = strTemp										
	End Function
	
	private Function BuildBottomBar()
	    Dim strTemp	
		if (s_showBottomBar) then
			strTemp = "<tr><td class='CuteEditorBottomBarContainer'>"
			strTemp = strTemp & "<table border='0' cellspacing='0' cellpading='0' style='width:100%;'><tr><td class='CuteEditorBottomBarContainer' style='width:150px'>"
			strTemp = strTemp & "<img alt="""&G_Str("Normal")&""" Command=""TabEdit""  src="""&ProcessThemeWebPath("design.gif")&""" border=""0"" />"
			if s_showhtmlmode = "true" then
				strTemp = strTemp & "<img alt="""&G_Str("HTML")&""" Command=""TabCode""  src="""&ProcessThemeWebPath("htmlview.gif")&""" border=""0"" />"
			end if
			if s_showpreviewmode = "true" then
				strTemp = strTemp & "<img alt="""&G_Str("Preview")&""" Command=""TabView""  src="""&ProcessThemeWebPath("preview.gif")&""" border=""0"" />"
			end if
			strTemp = strTemp & "</td>"
			if s_ShowTagSelector = "true" then
			    strTemp = strTemp & "<td><div id='"&ClientID&"_TagListContainer' class='CuteEditorTagListContainer'>&nbsp;</div></td>"
			end if
			if s_ShowWordCount= "true" then
			    strTemp = strTemp & "<td style='text-align:right;' nowrap='nowrap'><span class='WordCount'></span><span class='WordSpliter'>&nbsp;</span><span class='CharCount'></span></td>"
			end if
			strTemp = strTemp & "<td style='text-align:right;'>"
			if Lcase(s_ResizeMode) = "plusminus" then
			    if s_ShowEnlargeButton then
				    strTemp = strTemp & "<img alt="""&G_Str("Enlarge")&""" Command=""sizeplus"" src="""&ProcessThemeWebPath("plus.gif")&""" border=""0""/>"
			    end if
			    if s_ShowDecreaseButton then
				    strTemp = strTemp & "<img alt="""&G_Str("Decrease")&""" Command=""sizeminus"" src="""&ProcessThemeWebPath("minus.gif")&""" border=""0""/>"
			    end if
			elseif Lcase(s_ResizeMode) = "resizecorner" then
				    strTemp = strTemp & "<img alt="""&G_Str("Resize")&""" ondragstart=""return false"" onmouseover=""(CuteEditor_GetEditor(this).RegisterResizeCornor||function(){})(this.parentNode)"" src="""&GetURL("Images/ResizeCorner.gif")&""" border=""0""/>"
			end if 
			strTemp = strTemp & "</td>"
			strTemp = strTemp & "</tr></table>"
			strTemp = strTemp & "</td></tr>"
		end if
		BuildBottomBar = strTemp 
	End Function
		
	
	private Function EditorInitialise()
		dim t		
		dim loaderFolder
		Select case Lcase(BrowserType)
			Case "safari12":
				loaderFolder = "safari12_Loader"
			Case "safari":
				loaderFolder = "safari_Loader"
			Case "opera":
				loaderFolder = "opera_Loader"
			Case "gecko":
				loaderFolder = "gecko_Loader"
			Case "winie":
				loaderFolder = "ie_Loader"
		End Select
		t = "<script language=""JavaScript"">"
		
		t = t & "var CE_Editor1_IDSettingClass_Strings={"		
		t = t & GetAllStringByCulture (s_CustomCulture)	
		t = t & "'':''};"
		t = t & "</script>"
		
		
		t = t & "<script language=""JavaScript"" src="""&s_filespath&"/scripts/"&loaderFolder&"/Loader.js""></script>"
		
		Randomize
		
		t = t & "<img src="""&s_filespath&"/images/1x1.gif?"&Rnd&""""		
		t = t & " onload=""CuteEditorInitialize('"&ClientID&"',{"		
		t = t & "'_ClientID':'"&ClientID&"',"
		t = t & "'_UniqueID':'"&ID&"',"
		t = t & "'_FrameID':'"&ClientID&"_Frame',"
		t = t & "'_ToolBarID':'"&ClientID&"_ToolBar',"
		t = t & "'_CodeViewToolBarID':'"&ClientID&"_CodeViewToolBar',"
		t = t & "'_HiddenID':'"&ID&"',"
		t = t & "'_StateID':'"&ID&"$ClientState',"
		t = t & "'Culture':'"&s_CustomCulture&"',"
		t = t & "'Theme':'"&s_ThemeType&"',"
		t = t & "'ResourceDir':'"&s_filespath&"',"
		t = t & "'ActiveTab':'"&s_activetab&"',"
		t = t & "'ToggleBorder':'"&s_ToggleBorder&"',"
		t = t & "'FullPage':'"&s_FullPage&"',"
		if s_ContextMenuMode <> "" Then
		    t = t & "'ContextMenuMode':'"&s_ContextMenuMode&"',"
		else
		    t = t & "'ContextMenuMode':'"&s_autoconfigure&"',"
		End if
		t = t & "'EnableBrowserContextMenu':'"&s_EnableBrowserContextMenu&"',"
		t = t & "'EnableContextMenu':'"&s_EnableContextMenu&"',"
		t = t & "'FocusOnLoad':'"&s_Focus&"',"
		t = t & "'ConvertHTMLTagstoLowercase':'"&s_ConvertHTMLTagstoLowercase&"',"
		t = t & "'RemoveTBODYTag':'"&s_RemoveTBODYTag&"',"
		t = t & "'AllowEditServerSideCode':'"&s_AllowEditServerSideCode&"',"
		t = t & "'EnableAntiSpamEmailEncoder':'"&s_EnableAntiSpamEmailEncoder&"',"
		t = t & "'EnableStripScriptTags':'"&s_EnableStripScriptTags&"',"
		t = t & "'MaxHTMLLength':'"&s_MaxHTMLLength&"',"
		t = t & "'MaxTextLength':'"&s_MaxTextLength&"',"
		t = t & "'TabSpaces':'"&s_TabSpaces&"',"
		t = t & "'ResizeStep':'"&s_ResizeStep&"',"
		t = t & "'BreakElement':'"&s_breakelement&"',"
		t = t & "'ResizeMode':'"&s_ResizeMode&"',"
		t = t & "'URLType':'"&s_URLType&"',"
		t = t & "'EmptyAlternateText':'"&s_EmptyAlternateText&"',"
		t = t & "'HyperlinkTarget':'"&s_HyperlinkTarget&"',"
		t = t & "'ServerName':'"&s_ServerName&"',"
		t = t & "'AllowPasteHtml':'"&s_AllowPasteHtml&"',"
		t = t & "'EncodeHiddenValue':'False',"
		t = t & "'UseStandardDialog':'"&s_UseStandardDialog&"',"
		t = t & "'UseSimpleAmpersand':'"&s_UseSimpleAmpersand&"',"
		t = t & "'UseHTMLEntities':'"&s_UseHTMLEntities&"',"
		t = t & "'UsePhysicalFormattingTags':'"&s_UsePhysicalFormattingTags&"',"
		t = t & "'EnableObjectResizing':'"&s_EnableObjectResizing&"',"
		t = t & "'EnableDragDrop':'"&s_EnableDragDrop&"',"
		t = t & "'MaintainAspectRatioWhenDraggingImage':'"&s_MaintainAspectRatioWhenDraggingImage&"',"
		t = t & "'UseFontTags':'"&s_UseFontTags&"',"
		t = t & "'RenderRichDropDown':'"&RenderRichDropdown&"',"
		t = t & "'EditorOnPaste':'"&s_EditorOnPaste&"',"
		t = t & "'EditorWysiwygModeCss':'"&s_EditorWysiwygModeCss&"',"
		t = t & "'PreviewModeCss':'"&s_PreviewModeCss&"',"
		t = t & "'HelpPath':'"&s_HelpUrl&"',"
		t = t & "'PrintFullWebPage':'"&s_PrintFullWebPage&"',"
		t = t & "'ReadOnly':'"&s_ReadOnly&"',"
		t = t & "'EditCompleteDocument':'"&s_EditCompleteDocument&"',"
		t = t & "'DOCTYPE':'"&server.HTMLEncode(s_DOCTYPE)&"',"
		t = t & "'BaseHref':'"&s_BaseHref&"',"
		t = t & "'EditorBodyStyle':'"&s_EditorBodyStyle&"',"
		t = t & "'EditorBodyId':'"&s_EditorBodyId&"',"
		t = t & "'EditorBodyClass':'"&s_EditorBodyClass&"',"
		t = t & "'XHTMLOutput':'"&s_XMLOutput&"',"
		t = t & "'EditorSetting2':'"&EditorSetting2&"',"
		t = t & "'Theme':'"&s_ThemeType&"',"
		t = t & "'EditorSetting':'"&EditorSetting&"'})"""
		if Lcase(BrowserType) <> "opera" then
		    t = t & " style=""display:none;"""
		end if
		t = t & "/>"
		EditorInitialise = t	
	End Function
	
	private Function editorClientScript ()
	    Dim strTemp
		if s_subsequent = "false" then 
			strTemp = "<script language=""JavaScript"" SRC="""&s_filespath&"/Scripts/spell.js""></script><script language=""JavaScript"" SRC="""&s_filespath&"/Scripts/Constant.js""></script>"
		end if
		editorClientScript =  strTemp
	End Function
	
	private Function editorStylesheet ()
		editorStylesheet =  "<link href='"&s_filespath&"/Themes/"&s_ThemeType&"/style.asp?EditorID="&ClientID&"' type='text/css' rel='stylesheet'/>"
	End Function
	
	Public Sub Draw()
		Response.Write GetString()
	End Sub
	
	
	Public Function GetString()
		Dim s, strTemp			
		if Not s_enableclientscript OR BrowserType="false" then
			if s_AccessKey <> "" then
				s=s&" AccessKey="""&s_AccessKey&""""
			end if 
			strTemp = "<textarea name="""&ID&""" id='"&ID&"' "&s&" rows="""&s_DownLevelRows &""" cols="""&s_DownLevelColumns&""" style=""width: " & s_width & "; height: " & s_height & """ ID="""&ID&""">" &  s_Text & "</textarea>"
		Else	
			strTemp = strTemp & " <!-- CuteEditor Version 6.6 "&ID&" Begin --> " & vbCRLF
			if Lcase(BrowserType)="safari12" then
			    strTemp = "<input name="""&ID&""" id='"&ID&"' type=hidden value='" &  s_Text & "'>"
			else
			    strTemp = strTemp &  "<textarea name='"&ID&"' id='"&ID&"' rows='"&s_DownLevelRows &"' cols='"&s_DownLevelColumns&"' class='CuteEditorTextArea' style='DISPLAY: none; WIDTH: 100%; HEIGHT: 100%'>" & s_Text & "</textarea>"
			end if
			strTemp = strTemp & editorClientScript ()
			strTemp = strTemp & editorStylesheet ()
			strTemp = strTemp & ""
			if s_tabindex <> 0 then
				s="TabIndex="&s_tabindex
			end if 
			if s_AccessKey <> "" then
				s=s&" AccessKey="""&s_AccessKey&""""
			end if 
			Dim  start_time
			start_time = Timer
			strTemp = strTemp & "<input type=hidden name="""&ID&"$ClientState"" value=''/>"
			strTemp = strTemp & "<table "&s&" id="""&ClientID&""" _IsCuteEditor=""True"" cellspacing=""0"" cellpadding=""0"" height="""&s_height&""" width="""&s_width&""" style=""background-color:"&s_BackColor&";border-color:"&s_BorderColor&";border-width:"&s_BorderWidth&"px;border-style:"&s_BorderStyle&";table-layout:auto;height:"&s_height&"px;width:"&s_width&"px;"">"
			strTemp = strTemp & "<tr><td class=""CuteEditorToolBarContainer"" unselectable='on' valign='top'>"
			strTemp = strTemp & "<div id="""&ClientID&"_ToolBar"" style=""display:none;"">"
			if lcase(s_ShowToolBar) = "true" then
			    strTemp = strTemp & CreateToolBar()
			    CustomAddons=CustomAddons & ""
				' strTemp = strTemp &  CustomAddons
				strTemp = replace(strTemp, "##HOLDER##", CustomAddons) 
			end if
			strTemp = strTemp & "</div>"
			strTemp = strTemp & "<div id="""&ClientID&"_CodeViewToolBar"" style=""display:none;"">"		
				
			if lcase(s_ShowCodeViewToolBar) = "true" then
			    strTemp = strTemp & GetToolbarFromItemList(s_CodeViewTemplateItemList,s_disableitemlist)
			end if			
			
			strTemp = strTemp & "</div>"
			strTemp = strTemp & "</td></tr>"
			strTemp = strTemp & "<tr><td class=""CuteEditorFrameContainer"">"
			strTemp = strTemp & "<iframe id="""&ClientID&"_Frame"" src="""&GetURL("template.asp")&""" FrameBorder=""0"" class=""CuteEditorFrame"" style=""background-color:White;border-color:#dddddd;border-width:1px;border-style:Solid;height:100%;width:100%;""></iframe>"
			strTemp = strTemp & "</td></tr>"
			strTemp = strTemp & BuildBottomBar ()
			strTemp = strTemp & "</table>"
			strTemp = strTemp & EditorInitialise ()
			strTemp = strTemp & vbCRLF & " <!-- CuteEditor "&ID&" End "&Timer- start_time&"s--> " & vbCRLF
		end if
		GetString = strTemp
	End Function
	
	'--------------------------------------	
	
	Public Property Get EditorSetting
	    dim s
		s=""
		
		if CStr(s_maxImageSize) <> "" then
		    s=s&s_maxImageSize&"|"
		Else
		    s=s&G_SettingFromSecurityPolicyFile("MaxImageSize")&"|"
		end if
		
		if CStr(s_MaxMediaSize) <> "" then
		    s=s&s_MaxMediaSize&"|"
		Else
		    s=s&G_SettingFromSecurityPolicyFile("MaxMediaSize")&"|"
		end if
		
		if CStr(s_MaxFlashSize) <> "" then
		    s=s&s_MaxFlashSize&"|"
		Else
		    s=s&G_SettingFromSecurityPolicyFile("MaxFlashSize")&"|"
		end if
		
		if CStr(s_MaxDocumentSize) <> "" then
		    s=s&s_MaxDocumentSize&"|"
		Else
		    s=s&G_SettingFromSecurityPolicyFile("MaxDocumentSize")&"|"
		end if
		
		if CStr(s_MaxTemplateSize) <> "" then
		    s=s&s_MaxTemplateSize&"|"
		Else
		    s=s&G_SettingFromSecurityPolicyFile("MaxTemplateSize")&"|"
		end if
		
		if CStr(s_ImageGalleryPath) <> "" then
		    s=s&s_ImageGalleryPath&"|"
		Else
		    s=s&G_SettingFromSecurityPolicyFile("ImageGalleryPath")&"|"
		end if
		
		if CStr(s_MediaGalleryPath) <> "" then
		    s=s&s_MediaGalleryPath&"|"
		Else
		    s=s&G_SettingFromSecurityPolicyFile("MediaGalleryPath")&"|"
		end if
		
		if CStr(s_FlashGalleryPath) <> "" then
		    s=s&s_FlashGalleryPath&"|"
		Else
		    s=s&G_SettingFromSecurityPolicyFile("FlashGalleryPath")&"|"
		end if
		
		if CStr(s_TemplateGalleryPath) <> "" then
		    s=s&s_TemplateGalleryPath&"|"
		Else
		    s=s&G_SettingFromSecurityPolicyFile("TemplateGalleryPath")&"|"
		end if
		
		if CStr(s_FilesGalleryPath) <> "" then
		    s=s&s_FilesGalleryPath&"|"
		Else
		    s=s&G_SettingFromSecurityPolicyFile("FilesGalleryPath")&"|"
		end if
		
		if CStr(s_AllowUpload) <> "" then
		    s=s&s_AllowUpload&"|"
		Else
		    s=s&G_SettingFromSecurityPolicyFile("AllowUpload")&"|"
		end if
		
		if CStr(s_AllowCreateFolder) <> "" then
		    s=s&s_AllowCreateFolder&"|"
		Else
		    s=s&G_SettingFromSecurityPolicyFile("AllowCreateFolder")&"|"
		end if
		
		if CStr(s_AllowRename) <> "" then
		    s=s&s_AllowRename&"|"
		Else
		    s=s&G_SettingFromSecurityPolicyFile("AllowRename")&"|"
		end if
		if CStr(s_AllowDelete) <> "" then
		    s=s&s_AllowDelete&"|"
		Else
		    s=s&G_SettingFromSecurityPolicyFile("AllowDelete")&"|"
		end if
		if CStr(s_ImageFilters) <> "" then
		    s=s&s_ImageFilters&"|"
		Else
		    s=s&G_SettingFromSecurityPolicyFile("ImageFilters")&"|"
		end if
				
		if CStr(s_MediaFilters) <> "" then
		    s=s&s_MediaFilters&"|"
		Else
		    s=s&G_SettingFromSecurityPolicyFile("MediaFilters")&"|"
		end if
		
		if CStr(s_DocumentFilters) <> "" then
		    s=s&s_DocumentFilters&"|"
		Else
		    s=s&G_SettingFromSecurityPolicyFile("DocumentFilters")&"|"
		end if		
		if CStr(s_TemplateFilters) <> "" then
		    s=s&s_TemplateFilters&"|"
		Else
		    s=s&G_SettingFromSecurityPolicyFile("TemplateFilters")&"|"
		end if	
		s=s&s_CustomCulture&"|"		
		s=s&G_SettingFromSecurityPolicyFile("DemoMode")
	'	s=s&s_filespath&"|"	
		call initCodecs
	'	EditorSetting = server.URLEncode(Base64encode(s))
	    EditorSetting = base64Encode(s)
	    Response.Cookies("CESecurity")=EditorSetting
	    Session("CESecurity")=EditorSetting
	End Property	
	Public Property Get EditorSetting2
	    dim s
		s=""
		s=s&Request.ServerVariables("server_name")&"|"
		s=s&Request.ServerVariables("HTTP_HOST")&"|"
		s=s&Request.ServerVariables("LOCAL_ADDR")&"|"
		s=s&Request.ServerVariables("REMOTE_HOST")&"|"
		s=s&"("&year(date)&","&month(date)&","&day(date)&")"&"|"
	    EditorSetting2 = server.URLEncode(s)
	End Property
	Public Function G_SettingFromSecurityPolicyFile(instring)
		dim scriptname,xmlfilename,doc,temp
		dim node,selectednode,optionnodelist,errobj
		dim selectednodes,i,Nodes,objNode

		xmlfilename = Server.MapPath(GetURL("Configuration/Security/"&s_SecurityPolicyFile&""))

		' Create an object to hold the XML
		set doc = server.CreateObject("Microsoft.XMLDOM")

		' For ASP, wait until the XML is all ready before continuing
		doc.async = False

		' Load the XML file or return an error message and stop the script
		if not Doc.Load(xmlfilename) then
			Response.Write "Failed to load the language text from the XML file.<br>" & instring
			Response.End
		end if

		' Make sure that the interpreter knows that we are using XPath as our selection language
		doc.setProperty "SelectionLanguage", "XPath"
	
	    if InStr(1, instring, "Filters", 1) > 0 then
	        set Nodes = doc.DocumentElement.selectNodes("/configuration/security[@name='"&instring&"']/item")
	        dim s
		    For Each objNode in Nodes
			    s = s&objNode.Text&","
			Next 
	        G_SettingFromSecurityPolicyFile= s
	    Else 
		    set selectednode= doc.selectSingleNode("/configuration/security[@name='"&instring&"']")
		    if IsObject(selectednode) and not selectednode is nothing  then
			    G_SettingFromSecurityPolicyFile= Server.HTMLEncode(selectednode.text)
		    else
			    G_SettingFromSecurityPolicyFile= ""		
		    end if
		end if
	End Function
	
	public Sub LoadHTML(ByVal FilePath)
		dim fso
		dim file
		dim fileContents
			
		set fso = Server.CreateObject("Scripting.FileSystemObject")
		FilePath = Server.mapPath(FilePath)
			
		if fso.FileExists(FilePath) then
			set file = fso.OpenTextFile(FilePath, 1, "false", -2)
			if not (file.AtEndOfStream) then 
                fileContents = file.ReadAll
            end if
			Text = fileContents
			exit Sub
		else
			Text = "File " & FilePath & " doesn't exist"
		end if
	end Sub
	
	
	public Sub SaveFile (ByVal FilePath)
		dim fso
		dim file
		dim stream
		dim fileContents
			
		set fso = Server.CreateObject("Scripting.FileSystemObject")
		FilePath = Server.MapPath(FilePath)
			
		if len(Text) = 0 then
			response.Write ""
			exit Sub
		else
			if NOT fso.FileExists(FilePath) then
				set file = fso.CreateTextFile(FilePath)
			end if
			set file = fso.GetFile(FilePath)
			set stream = file.OpenAsTextStream(2)					
			stream.Write(Text)
			stream.Close
		end if
	end Sub
	
	private Property Get SaveButton	
		SaveButton = "<INPUT type=""image"" src="""&ProcessThemeWebPath("save.gif")&""" name=""Save"" title="""&G_Str("Save")&""" class=""CuteEditorButton"">"
	End Property
		
	private Property Get GetURL(path)
		GetURL = ""&s_filespath&"/"&path	
	End Property	
	
	private Property Get ProcessThemeWebPath(imageURL)
		Select case Lcase(s_ThemeType)
			Case "office2007":
				ProcessThemeWebPath = GetURL("Themes/Office2007/Images/"&imageURL)
			Case "Office2003Blue":
				ProcessThemeWebPath = GetURL("Themes/Office2003Blue/Images/"&imageURL)
			Case "office2003":
				ProcessThemeWebPath = GetURL("Themes/office2003/Images/"&imageURL)
			Case "officexp":
				ProcessThemeWebPath = GetURL("Themes/officexp/Images/"&imageURL)
			Case "office2000":
				ProcessThemeWebPath = GetURL("Themes/office2000/Images/"&imageURL)		
			Case Else
				ProcessThemeWebPath = GetURL("Themes/office2007/Images/"&imageURL)
		End Select
	End Property
	
	Public Function G_Str(instring)
		dim t		
		t = GetStringByCulture(instring,s_CustomCulture)
        		
		if t = ""  then
			t= GetStringByCulture(instring,"_default")
		end if
		
		if t = ""  then
			t= "{"&instring&"}"	
		end if
		G_Str= t
	End Function	
	
	Public Function GetAllStringByCulture(input_culture)
		dim scriptname,xmlfilename,doc,temp
		dim node,Nodes,objNode,objName
		
		xmlfilename= Server.MapPath(GetURL("languages/"&input_culture&".xml"))

		' Create an object to hold the XML
		set doc = server.CreateObject("Microsoft.XMLDOM")

		' For ASP, wait until the XML is all ready before continuing
		doc.async = False
		Doc.Load(xmlfilename)	
        set Nodes = doc.DocumentElement.selectNodes("//resources/*")                    
        dim t
        t=""        
        For Each objNode in Nodes
            With objNode.Attributes
                set objName = .GetNamedItem("name") 
                t=t&""""&lcase(objName.Text)&""":"""&objNode.Text&""","
			End With          
		Next                    
		GetAllStringByCulture=t		
	End Function
	
	private xmldocs
	
	public function getxmldoc(name)
	
		dim xmlfilename,doc	
	
		if not xmldocs.Exists (name) then
			xmlfilename= Server.MapPath(GetURL("languages/"&name&".xml"))

			' Create an object to hold the XML
			set doc = server.CreateObject("Microsoft.XMLDOM")

			' For ASP, wait until the XML is all ready before continuing
			doc.async = False

			' Load the XML file or return an error message and stop the script
			if not Doc.Load(xmlfilename) then
			Response.Write "Failed to load the language text from the XML file.<br>" & instring
			Response.End
			end if

			' Make sure that the interpreter knows that we are using XPath as our selection language
			doc.setProperty "SelectionLanguage", "XPath"
			xmldocs.Add name,doc
			
		end if
		
		set getxmldoc=xmldocs(name)
	
	
	end function
	
	Public Function GetStringByCulture(instring,input_culture)
	
		dim selectednode, selectednodes, doc
		
		set doc=getxmldoc(input_culture)
	
		set selectednode= doc.selectSingleNode("/resources/resource[@name='"&instring&"']")
		if IsObject(selectednode) and not selectednode is nothing  then
			GetStringByCulture=Server.HTMLEncode(selectednode.text)
		else
			GetStringByCulture=""		
		end if
		
	End Function
		
	Public function GetToolbarFromItemList(templatelist,d_list)
		dim scriptname,cfgfilename,doc,temp
		dim node,selectednode,optionnodelist,errobj
		dim selectednodes,s
		s=""
				
		cfgfilename = Server.MapPath(GetURL("Configuration/AutoConfigure/full.config"))
		
		' Create an object to hold the XML
		set doc = server.CreateObject("Microsoft.XMLDOM")

		' For ASP, wait until the XML is all ready before continuing
		doc.async = False

		' Load the XML file or return an error message and stop the script
		if not Doc.Load(cfgfilename) then
			Response.Write "Failed to load the Configure file.<br>"
			Response.End
		end if

		' Make sure that the interpreter knows that we are using XPath as our selection language
		doc.setProperty "SelectionLanguage", "XPath"
	
		Dim Nodes
		set Nodes = doc.DocumentElement.selectNodes("//toolbars/*")
		
		Dim ToolbarStrings
		Dim i
				
		If	Lcase(BrowserType) <> "winie" then
			d_list = d_list & ",zoom"
		end if
		
		If	BrowserType="Safari12" then		
			d_list = d_list & ",zoom,find,insertorderedlist,insertTemplate,insertunorderedlist,indent,outdent,imagemap,strikethrough"
		end if
		
		ToolbarStrings = Split(templatelist,",")
		For i = 0 to Ubound(Split(templatelist,","))
			dim itemname			
			itemname = Trim(ToolbarStrings(i))
			
			dim disable_toolbarstrings, j,found, objCommand,objWidth
			found = false
			disable_toolbarstrings = Split(d_list,",")
			for j = 0 to Ubound(disable_toolbarstrings) 					
				if lcase(itemname) = lcase(trim(disable_toolbarstrings(j))) then
					found = true
					Exit for					
				end if
			next
			if found = false then
			Select Case lcase(itemname)
				case "g_start":
					s =  s & 	AddToolbarGroupStart
				case "g_end":
					s =  s & 	AddToolbarGroupEnd
				case "separator":
					s =  s & 	AddToolbarSeparator
				case "linebreak":
					s =  s & 	AddToolbarLineBreak
				case "table":
					s =  s & 	AddToolbarTable
				case "forecolor":
					s =  s & 	AddToolbarForeColor
				case "backcolor":
					s =  s & 	AddToolbarBackColor
				case "dropdown":
			'		s =  s & 	AddToolbarLineBreak
				case "holder":
					s = s & "##HOLDER##"
				Case else
					Dim objNode,objType,objName,objImgName,objVisible,v
					For Each objNode in Nodes
						With objNode.Attributes	
							set objType = .GetNamedItem("type") 
				            set objName = .GetNamedItem("name") 
				            set objCommand = .GetNamedItem("command")
				            set objVisible = .GetNamedItem("visible")
							set objImgName = .GetNamedItem("imagename") 
							set objWidth = .GetNamedItem("width") 
				
							v = "true"													
							if not objVisible is nothing  then
								v = objVisible.Text
							end if
							if not objName is nothing AND lcase(v) <> "false" then							    
							    if lcase(objName.Text) = lcase(itemname) then	
							        Select Case lcase(objType.Text)
			                            case "image":
			                                dim n,t,c,w							
										    n = objName.Text	
										    if not objCommand is nothing  then
											    c = objCommand.Text
										    else
											    c = n								
										    end if 
										    if not objImgName is nothing  then
											    t = objImgName.Text
										    else
											    t = n								
										    end if 	
										    s =  s & AddToolbarItem(n,c,t,20,20)									
			                            case "dropdown":					
										    n = objName.Text	
										    if not objCommand is nothing  then
											    c = objCommand.Text
										    else
											    c = n								
										    end if 
										    w="40"										
										    if not objWidth is nothing  then
											    w = objWidth.Text
										    end if 
										    s =  s & AddToolbarDropDown(n,c,w)
		                            End Select
		                        end if 
							end if	
						End With						
					Next										
			End Select
			end if
		Next	
		GetToolbarFromItemList = s
	End function
	
	Public function GetToolbarItems(cfgfilename,d_list)
		dim scriptname,doc,temp
		dim node,selectednode,optionnodelist,errobj
		dim selectednodes,i,s,Nodes,objNode,objType,objName,objImgName,objCommand,objWidth
		s=""
				
		cfgfilename = Server.MapPath(cfgfilename)
		
		' Create an object to hold the XML
		set doc = server.CreateObject("Microsoft.XMLDOM")

		' For ASP, wait until the XML is all ready before continuing
		doc.async = False

		' Load the XML file or return an error message and stop the script
		if not Doc.Load(cfgfilename) then
			Response.Write "Failed to load the Configure file.<br>"
			Response.End
		end if
		
		If	Lcase(BrowserType) <> "winie" then
			d_list = d_list & ",zoom"
		end if
		
		If	BrowserType="Safari12" then		
			d_list = d_list & ",zoom,find,insertorderedlist,insertunorderedlist,indent,outdent,imagemap,strikethrough"
		end if		

		' Make sure that the interpreter knows that we are using XPath as our selection language
		doc.setProperty "SelectionLanguage", "XPath"
	
		'set selectednode= doc.selectSingleNode("/toolbars")
		set Nodes = doc.DocumentElement.selectNodes("//toolbars/*")
		
		For Each objNode in Nodes
			With objNode.Attributes
				Dim objVisible, v
				set objType = .GetNamedItem("type") 
				set objName = .GetNamedItem("name") 
				set objCommand = .GetNamedItem("command")
				set objVisible = .GetNamedItem("visible")
				v = "true"							
				if not objVisible is nothing  then
					v = objVisible.Text
				end if
				if not objType is nothing AND lcase(v) <> "false" then
					dim disable_toolbarstrings, j,found
					found = "false"
					disable_toolbarstrings = Split(d_list,",")
					for j = 0 to Ubound(disable_toolbarstrings) 					
						if (lcase(trim(objType.Text)) = lcase(trim(disable_toolbarstrings(j)))) then
							found = "true"
							Exit for
						else
							if not objName is nothing  then
								if (lcase(trim(.GetNamedItem("name").Text)) = lcase(trim(disable_toolbarstrings(j)))) then
									found = "true"
									Exit for
								end if
							end if							
						end if
					next
					
					dim n,c,w,t					
					if(found = "false") then					
						Select Case lcase(objType.Text)
							case "g_start":
								s =  s & 	AddToolbarGroupStart
							case "g_end":
								s =  s & 	AddToolbarGroupEnd
							case "separator":
								s =  s & 	AddToolbarSeparator
							case "linebreak":
								s =  s & 	AddToolbarLineBreak
							case "table":
								s =  s & 	AddToolbarTable
							case "forecolor":
								s =  s & 	AddToolbarForeColor
							case "backcolor":
								s =  s & 	AddToolbarBackColor
							case "dropdown":
							    if not objName is nothing  then		
									n=.GetNamedItem("name").Text	
									c="PasteHTML"	
									if not objCommand is nothing then
									    c = .GetNamedItem("command").Text	
									end if		
									w="40"	
									set objWidth = .GetNamedItem("width") 
									if not objWidth is nothing  then
										w = .GetNamedItem("width").Text
									end if						
									s =  s & AddToolbarDropDown(n,c,w)
								end if
							case "holder":
								s = s & "##HOLDER##"
							Case "image"
								if not objName is nothing  then		
								    n = .GetNamedItem("name").Text									    
									if not objCommand is nothing then
									    c = .GetNamedItem("command").Text									
									else									    
									    c = n	
									end if									
									set objImgName = .GetNamedItem("imagename") 
									if not objImgName is nothing  then
										t = .GetNamedItem("imagename").Text
									else
										t = n								
									end if 	
									s =  s & AddToolbarItem(n,c,t,20,20)						
								end if				
							Case else
								s =  s & AddToolbarSeparator
						End Select
					end if
				end if				
			End With
		Next 
		GetToolbarItems = s
	End function
	
	
	Public function AddToolbarDropDownfromConfig(name,command,width)
		dim scriptname,cfgfilename,doc,temp,Nodes,objNode,objText,objValue,objHTML
		dim node,selectednode,optionnodelist,errobj
		dim selectednodes,i,s
		s=""
				
		cfgfilename = Server.MapPath(GetURL("Configuration/Shared/Common.config"))
		
		' Create an object to hold the XML
		set doc = server.CreateObject("Microsoft.XMLDOM")

		' For ASP, wait until the XML is all ready before continuing
		doc.async = False

		' Load the XML file or return an error message and stop the script
		if not Doc.Load(cfgfilename) then
			Response.Write "Failed to load the Configure file.<br>"
			Response.End
		end if

		' Make sure that the interpreter knows that we are using XPath as our selection language
		doc.setProperty "SelectionLanguage", "XPath"
	
		'set selectednode= doc.selectSingleNode("/toolbars")
		set Nodes = doc.DocumentElement.selectNodes("//dropdowns/"&name&"/*")
		
		if lcase(RenderRichDropdown) <> "true" or BrowserType="safari12" or BrowserType="safari" then		
		    s = "<select id="""&ClientID&lcase(name)&"""  OnChange=""CuteEditor_DropDownCommand(this,'"&command&"')"" class=""CuteEditorSelect"">"
		    s = s&"<option value=''>"&G_Str(name)&"</option>"
		    For Each objNode in Nodes
			    With objNode.Attributes
				    set objText = .GetNamedItem("text") 
				    if not objText is nothing then		
				        t=objText.Text 
				        v=t
				        h=t
				        o=t
                        if left (t,2) = "[["  then
                            t = Mid(t,3,len(t)-4)
				            t=G_Str(t)
                        end if                  
                        
				        set objValue = .GetNamedItem("value")
		                if  objValue is nothing  then
				            set objValue= objNode.selectSingleNode("value")
		                end if					        
				        if not objValue is nothing then
				            v=objValue.Text
				        end if	
				        v=Server.HTMLEncode(v)	
				        
				        set objHTML = .GetNamedItem("html")
		                if  objHTML is nothing  then
				            set objHTML= objNode.selectSingleNode("html")
		                end if	 
		                
		                if IsObject(objHTML) and not objHTML is nothing  then
			                h=objHTML.text
		                end if	 
		                h=Replace(h,o,t)
		           '    h=Server.HTMLEncode(h)			    
					    s = s&"<option value='"&v&"'>"&t&"</option>"
				    end if				
			    End With
		    Next 
		    s = s&"</select>"
		else		
		    s = "<span class='CuteEditorDropDown' id='"&ClientID&lcase(name)&"' Command='"&command&"' "
		    s = s&"onchange=""CuteEditor_DropDownCommand(this,'"&command&"')"" RichHideFirstItem='true' _IsRichDropDown='True' style='display:inline-block;width:"&width&"px;height:20px;'>"
		    s = s&"<span val="""" selected='True' html='"&G_Str(name)&"' txt="""&G_Str(name)&"""></span>"
		     For Each objNode in Nodes
			    With objNode.Attributes
				    set objText = .GetNamedItem("text") 
				    if not objText is nothing  then				    
				        dim t, h , v, o
				        t=objText.Text 
				        v=t
				        h=t
				        o=t
                        if left (t,2) = "[["  then
                            t = Mid(t,3,len(t)-4)
				            t=G_Str(t)
                        end if                  
                        
				        set objValue = .GetNamedItem("value")
		                if  objValue is nothing  then
				            set objValue= objNode.selectSingleNode("value")
		                end if					        
				        if not objValue is nothing then
				            v=objValue.Text
				        end if	
				        v=Server.HTMLEncode(v)	
				        
				        set objHTML = .GetNamedItem("html")
		                if  objHTML is nothing  then
				            set objHTML= objNode.selectSingleNode("html")
		                end if	 
		                
		                if IsObject(objHTML) and not objHTML is nothing  then
			                h=objHTML.text
		                end if	 
		                h=Replace(h,o,t)
		           '    h=Server.HTMLEncode(h)		        		                		    
					    s = s&"<span val='"&v&"' html="""&h&""" txt='"&t&"'></span>"
				    end if				
			    End With
		    Next 
		    s = s&"</span>"
		end if 
		AddToolbarDropDownfromConfig = s
	End function
		
	private Property Get AddToolbarGroupStart
		AddToolbarGroupStart = "<table class='CuteEditorGroupMenu' cellSpacing='0' cellPadding='0' border='0'><tr><td class='CuteEditorGroupMenuCell'><nobr>"
	End Property
	private Property Get AddToolbarGroupEnd
		AddToolbarGroupEnd = "</nobr></td></tr></table>"
	End Property	
	private Property Get AddToolbarSeparator
		AddToolbarSeparator = "<img src='"&ProcessThemeWebPath("separator.gif")&"' unselectable='on' class='separator'/>"
	End Property	
	private Property Get AddToolbarLineBreak
		AddToolbarLineBreak = "<br clear='both'/>"
	End Property	
	Public Function AddToolbarItem(name, command, img, l_width,l_height)
		if lcase(command) = "save" then
			AddToolbarItem = SaveButton 
		Else
			if IsNumeric(l_width) and IsNumeric(l_height) then	
				dim t, alt
				t="<img"
				
				if lcase(name) = "tofullpage" then
					t = "<img id='cmd_tofullpage'"
				end if			
				
				if lcase(name) = "fromfullpage" then
					t = "<img id='cmd_fromfullpage'"			
				end if
				if Lcase(BrowserType) = "winie" then
				    alt = "alt="""&G_Str(name)&""""
				else				    
				    alt = "alt="""&G_Str(name)&""""
				end if
				if lcase(command) = "syntaxhighlighter" then				
				    AddToolbarItem = t & " "&alt&" Command="""&command&""" src='"&ProcessThemeWebPath("code.gif")&"' />"
				else				
				    AddToolbarItem = t & " "&alt&" Command="""&command&""" ThemeIndex="""&GetImageThemeIndex(img)&""" />"
				end if
			end if
		end if
	End Function	
	
	
	Public Function AddToolbarDropDown(name,command, width)
		Dim s,i,name_array,list_array,name_s,list_s
		s = ""
		
		Select Case lcase(name)
			case "cssclass":
			    name_s=s_cssclassdropdownMenuNames
		        list_s=s_cssclassdropdownMenuList
			case "cssstyle":
			    name_s=s_inlinestyledropdownMenuNames
		        list_s=s_inlinestyledropdownMenuList
			case "formatblock":
			    name_s=s_ParagraphsListMenuNames
		        list_s=s_ParagraphsListMenuList
			case "fontname":
			    name_s=s_FontFacesList
		        list_s=s_FontFacesList
			case "fontsize":
			    name_s=s_FontSizesList
		        list_s=s_FontSizesList
			case "links":
			    name_s=s_linksdropdownMenuNames
		        list_s=s_linksdropdownMenuList
			case "codes":
			    name_s=s_codesnippetdropdownMenuNames
		        list_s=s_codesnippetdropdownMenuList
			case "images":
			    name_s=s_imagesdropdownMenuNames
		        list_s=s_imagesdropdownMenuList
			case "zoom":
			    name_s=s_ZoomsList
		        list_s=s_ZoomsList
			Case else
				'
		End Select
		dim l_value
		if name_s <> "" AND list_s <> "" then
		
            name_s = replace(name_s,"'","&#39;") 
            list_s = replace(list_s,"'","&#39;") 
            
		    name_array = Split(name_s,",")
		    list_array = Split(list_s,",")		    
		    if lcase(RenderRichDropdown) <> "true" or BrowserType="safari12" or BrowserType="safari" then		
				if  Lcase(BrowserType) = "gecko" then
					s = "<select id="""&ClientID&lcase(name)&"""  OnClick=""CuteEditor_DropDownCommand(this,'"&command&"')"" class=""CuteEditorSelect"">"
				else	
					s = "<select id="""&ClientID&lcase(name)&"""  OnChange=""CuteEditor_DropDownCommand(this,'"&command&"')"" class=""CuteEditorSelect"">"
				end if
		        s = s&"<option value=''>"&G_Str(name)&"</option>"		    
		        if IsArray(name_array) and IsArray(list_array) then
			        if not IsArrayEmpty(list_array) then
				        For i=0 to Ubound(list_array)							
							l_value=trim(list_array(i))
							if lcase(name)="images" then
								l_value="<img src="""+l_value+""" border=""0"" />"
							end if
					        s = s&"<option value='"&server.HTMLEncode(l_value)&"'>"&server.HTMLEncode(trim(name_array(i)))&"</option>"
				        Next
			        end if
		        end if 
		        s = s&"</select>"
		    else	 
		        s = "<span class='CuteEditorDropDown' id='"&ClientID&lcase(name)&"' Command='"&command&"' "
		        s = s&"onchange=""CuteEditor_DropDownCommand(this,'"&command&"')"" RichHideFirstItem='true' _IsRichDropDown='True' style='display:inline-block;width:"&width&"px;height:20px;'>"
		        s = s&"<span val="""" selected='True' html='"&G_Str(name)&"' txt="""&G_Str(name)&"""></span>"	    
		        if IsArray(name_array) and IsArray(list_array) then
			        if not IsArrayEmpty(list_array) then
				        For i=0 to Ubound(list_array)   					
							l_value=trim(list_array(i))
							if lcase(name)="images" then
								l_value="<img src="""+l_value+""" border=""0"" />"
							end if	
					        s = s&"<span val='"&server.HTMLEncode(l_value)&"' html='"&trim(name_array(i))&"' txt='"&trim(name_array(i))&"'></span>"
				        Next
			        end if
		        end if 
		        s = s&"</span>"		
		    end if
    		AddToolbarDropDown = s		
    	else
    		AddToolbarDropDown = AddToolbarDropDownfromConfig(name,command,width)    	    
		end if		
	End Function	
	private Property Get AddToolbarForeColor
		dim t,c
		t=""
		c=""&ClientID&"_forecolorimg"
		if s_ReadOnly <> "true" then
            if Lcase(BrowserType) = "safari12" then
                t="onmousedown=""CuteEditor_GetEditor(this).ExecCommand('ForeColor',false, document.getElementById('"&c&"').style.backgroundColor)"""  
            else
                t="onclick=""CuteEditor_GetEditor(this).ExecCommand('ForeColor',false, document.getElementById('"&c&"').style.backgroundColor,this)"""  
            end if	     
	    end if
	    AddToolbarForeColor = "<img id='"&c&"' Command=""ForeColor"" alt="""&G_Str("ForeColor")&""" src='"&ProcessThemeWebPath("fontcolor.gif")&"' width='17' height='20' border=0 style='background-color: red;' "&t&"/>"
	    t=""	    
		if s_ReadOnly <> "true" then
		    if Lcase(BrowserType) = "safari12" then
		       t="onmousedown=""CuteEditor_GetEditor(this).ExecCommand('SetForeColor',false,"&c&");"""
            else
               t="onclick=""CuteEditor_GetEditor(this).ExecImageCommand('SetForeColor',false,null,this);"" oncolorchange=""CuteEditor_GetEditor(this).ExecImageCommand('ForeColor',false,this.selectedColor,this); document.getElementById('"&c&"').style.backgroundColor = this.selectedColor"""
		    end if		
		end if
	    AddToolbarForeColor = AddToolbarForeColor & "<img Command=""SetForeColor"" alt="""&G_Str("SetForeColor")&""" src='"&ProcessThemeWebPath("tbdown.gif")&"' width='9' height='20' border='0' "&t&"/>"	    
	End Property		
	private Property Get AddToolbarBackColor
		dim t,c
		t=""
		c=""&ClientID&"_bkcolorimg"
		if s_ReadOnly <> "true" then
            if Lcase(BrowserType) = "safari12" then
                t="onmousedown=""CuteEditor_GetEditor(this).ExecCommand('BackColor',false, document.getElementById('"&c&"').style.backgroundColor)"""  
            else
                t="onclick=""CuteEditor_GetEditor(this).ExecCommand('BackColor',false, document.getElementById('"&c&"').style.backgroundColor,this)"""  
            end if	     
	    end if
	    AddToolbarBackColor = "<img id='"&c&"' Command=""BackColor"" alt="""&G_Str("BackColor")&""" src='"&ProcessThemeWebPath("colorpen.gif")&"' width='17' height='20' border=0 style='background-color: yellow;' "&t&"/>"
	    t=""	    
		if s_ReadOnly <> "true" then
		    if Lcase(BrowserType) = "safari12" then
		       t="onmousedown=""CuteEditor_GetEditor(this).ExecCommand('SetBackColor',false,"&c&");"""
            else
               t="onclick=""CuteEditor_GetEditor(this).ExecImageCommand('SetBackColor',false,null,this);"" oncolorchange=""CuteEditor_GetEditor(this).ExecImageCommand('BackColor',false,this.selectedColor,this); document.getElementById('"&c&"').style.backgroundColor = this.selectedColor"""
		    end if		
		end if
	    AddToolbarBackColor = AddToolbarBackColor & "<img Command=""SetBackColor"" alt="""&G_Str("SetBackColor")&""" src='"&ProcessThemeWebPath("tbdown.gif")&"' width='9' height='20' border='0' "&t&"/>"	    
	End Property		
	private Property Get AddToolbarTable
		dim t
		t=""
		if s_ReadOnly <> "true" then
            if Lcase(BrowserType) = "safari12" then
                t="onmousedown=""var editor=CuteEditor_GetEditor(this);editor.TableDropDown(this)"""  
            else
                t="onclick=""var editor=CuteEditor_GetEditor(this);editor.TableDropDown(this)"""  
            end if	     
	    end if
	    AddToolbarTable = "<img Command=""TableDropDown"" alt="""&G_Str("TableDropDown")&""" src='"&ProcessThemeWebPath("instable.gif")&"' "&t&"/>"
	End Property	
	
	Sub GetAllIndexMap ()		 
	    set d=Server.CreateObject("Scripting.Dictionary")			
			d.Add "save","0"
            d.Add "newdoc","1"
            d.Add "print","2"
            d.Add "bspreview","3"
            d.Add "find","4"
            d.Add "fit","5"
            d.Add "restore","6"
            d.Add "cleanup","7"
            d.Add "spell","8"
            d.Add "cut","9"
            d.Add "copy","10"
            d.Add "paste","11"
            d.Add "pastetext","12"
            d.Add "pasteword","13"
            d.Add "pasteashtml","14"
            d.Add "delete","15"
            d.Add "undo","16"
            d.Add "redo","17"
            d.Add "insertpagebreak","18"
            d.Add "insertdate","19"
            d.Add "timer","20"
            d.Add "specialchar","21"
            d.Add "emotion","22"
            d.Add "keyboard","23"
            d.Add "box","24"
            d.Add "layer","25"
            d.Add "groupbox","26"
            d.Add "image","27"
            d.Add "eximage","28"
            d.Add "flash","29"
            d.Add "media","30"
            d.Add "document","31"
            d.Add "template","32"
            d.Add "youtube","33"
            d.Add "insrow_t","34"
            d.Add "insrow_b","35"
            d.Add "delrow","36"
            d.Add "inscol_l","37"
            d.Add "inscol_r","38"
            d.Add "delcol","39"
            d.Add "inscell","40"
            d.Add "delcell","41"
            d.Add "row","42"
            d.Add "cell","43"
            d.Add "mrgcell_r","44"
            d.Add "mrgcell_b","45"
            d.Add "spltcell_r","46"
            d.Add "spltcell_b","47"
            d.Add "break","48"
            d.Add "paragraph","49"
            d.Add "left_to_right","50"
            d.Add "right_to_left","51"
            d.Add "form","52"
            d.Add "textarea","53"
            d.Add "textbox","54"
            d.Add "passwordfield","55"
            d.Add "hiddenfield","56"
            d.Add "listbox","57"
            d.Add "dropdownbox","58"
            d.Add "optionbutton","59"
            d.Add "checkbox","60"
            d.Add "imagebutton","61"
            d.Add "submit","62"
            d.Add "reset","63"
            d.Add "pushbutton","64"
            d.Add "page","65"
            d.Add "bold","66"
            d.Add "italic","67"
            d.Add "under","68"
            d.Add "left","69"
            d.Add "center","70"
            d.Add "right","71"
            d.Add "justifyfull","72"
            d.Add "justifynone","73"
            d.Add "unformat","74"
            d.Add "numlist","75"
            d.Add "bullist","76"
            d.Add "indent","77"
            d.Add "outdent","78"
            d.Add "superscript","79"
            d.Add "subscript","80"
            d.Add "strike","81"
            d.Add "ucase","82"
            d.Add "lcase","83"
            d.Add "rule","84"
            d.Add "link","85"
            d.Add "unlink","86"
            d.Add "anchor","87"
            d.Add "imagemap","88"
            d.Add "abspos","89"
            d.Add "forward","90"
            d.Add "backward","91"
            d.Add "borders","92"
            d.Add "selectall","93"
            d.Add "selectnone","94"
            d.Add "help","95"
	End Sub
    private Function GetImageThemeIndex(cmdname)            
         GetImageThemeIndex = d.item(Lcase(cmdname))
    End Function

End Class

const BASE_64_MAP_INIT = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"
     dim nl
     ' zero based arrays
     dim Base64EncMap(63)
     dim Base64DecMap(127)

     ' must be called before using anything else
     PUBLIC SUB initCodecs()
          ' init vars
          nl = "<P>" & chr(13) & chr(10)
          ' setup base 64
          dim max, idx
             max = len(BASE_64_MAP_INIT)
          for idx = 0 to max - 1
               ' one based string
               Base64EncMap(idx) = mid(BASE_64_MAP_INIT, idx + 1, 1)
          next
          for idx = 0 to max - 1
               Base64DecMap(ASC(Base64EncMap(idx))) = idx
          next
     END SUB

     ' encode base 64 encoded string
     PUBLIC FUNCTION base64Encode(plain)

          if len(plain) = 0 then
               base64Encode = ""
               exit function
          end if

          dim ret, ndx, by3, first, second, third
          by3 = (len(plain) \ 3) * 3
          ndx = 1
          do while ndx <= by3
               first  = asc(mid(plain, ndx+0, 1))
               second = asc(mid(plain, ndx+1, 1))
               third  = asc(mid(plain, ndx+2, 1))
               ret = ret & Base64EncMap(  (first \ 4) AND 63 )
               ret = ret & Base64EncMap( ((first * 16) AND 48) + ((second \ 16) AND 15 ) )
               ret = ret & Base64EncMap( ((second * 4) AND 60) + ((third \ 64) AND 3 ) )
               ret = ret & Base64EncMap( third AND 63)
               ndx = ndx + 3
          loop
          ' check for stragglers
          if by3 < len(plain) then
               first  = asc(mid(plain, ndx+0, 1))
               ret = ret & Base64EncMap(  (first \ 4) AND 63 )
               if (len(plain) MOD 3 ) = 2 then
                    second = asc(mid(plain, ndx+1, 1))
                    ret = ret & Base64EncMap( ((first * 16) AND 48) +((second \16) AND 15 ) )
                    ret = ret & Base64EncMap( ((second * 4) AND 60) )
               else
                    ret = ret & Base64EncMap( (first * 16) AND 48)
                    ret = ret & "="
               end if
               ret = ret & "="
          end if

          base64Encode = ret
     END FUNCTION


private Function IsArrayEmpty(varArray)
   Dim lngUBound
   On Error Resume Next
   lngUBound = UBound(varArray)
   if Err.Number <> 0 then
      IsArrayEmpty = True
   Else
      IsArrayEmpty = False
   end if
End Function

' Version 6.0 functions


%>