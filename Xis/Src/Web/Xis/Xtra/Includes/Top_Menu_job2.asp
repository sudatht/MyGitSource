	<div class="contentMenu">
		<table cellpadding="0" cellspacing="0" width="96%" ID="Table2">
			<tr>
				<td>
					<table cellpadding="0" cellspacing="2" ID="Table7">
						<tr>
							<% 
							if  (AToolbarAtrib(0,0)="1") then						
								strClass = "menu"
								strALinkStart = AToolbarAtrib(0,1)
								strALinkEnd = AToolbarAtrib(0,2)
								strJSEvents = "onMouseOver='menuOver(this.id);' onMouseOut='menuOut(this.id);'"							
							else
								strALinkStart = ""
								strALinkEnd = ""
								strClass = "menu disabled"
								strJSEvents = ""							
							end if
							%>						
							<td id="menu1" class="<%=strClass%>" <%=strJSEvents%>>
								<%=strALinkStart%>
								<img src="/xtra/images/icon_save.gif" width="18" height="15" alt="" align="absmiddle">Lagre
								<%=strALinkEnd%>								
							</td>				
							<% 
							if (lngOppdragID > 0) and (AToolbarAtrib(1,0) = "1") then						
								strClass = "menu"
								strALinkStart = AToolbarAtrib(1,1)
								strALinkEnd = AToolbarAtrib(1,2)
								strJSEvents = "onMouseOver='menuOver(this.id);' onMouseOut='menuOut(this.id);'"							
							else
								strALinkStart = ""
								strALinkEnd = ""
								strClass = "menu disabled"
								strJSEvents = ""							
							end if
							%>																
							<td id="menu2" class="<%=strClass%>" <%=strJSEvents%>>
								<%=strALinkStart%>
								<img src="/xtra/images/icon_job.gif" width="18" height="15" alt="" align="absmiddle">Vis
								<%=strALinkEnd%>
							</td>
							<% 
							if (HasUserRight(ACCESS_TASK, RIGHT_WRITE) = true) and (lngOppdragID > 0) and (AToolbarAtrib(2,0)="1")  then						
								strClass = "menu"
								strALinkStart = AToolbarAtrib(2,1)
								strALinkEnd = AToolbarAtrib(2,2)
								strJSEvents = "onMouseOver='menuOver(this.id);' onMouseOut='menuOut(this.id);'"							
							else
								strALinkStart = ""
								strALinkEnd = ""
								strClass = "menu disabled"
								strJSEvents = ""							
							end if
							%>																
							<td id="menu3" class="<%=strClass%>" <%=strJSEvents%>>
								<%=strALinkStart%>
								<img src="/xtra/images/icon_changeInfo.gif" width="18" height="15" alt="" align="absmiddle">Endre
								<%=strALinkEnd%>
							</td>
							<!-- Copy Commission -->
							<%
							if (HasUserRight(ACCESS_TASK, RIGHT_WRITE) = true) and (AToolbarAtrib(6,0)="1")  then
								strClass = "menu"
								strALinkStart = AToolbarAtrib(6,1)
								strALinkEnd = AToolbarAtrib(6,2)
								strJSEvents = "onMouseOver='menuOver(this.id);' onMouseOut='menuOut(this.id);'"
							else
								strALinkStart = ""
								strALinkEnd = ""
								strClass = "menu disabled"
								strJSEvents = ""
							end if
							%>
							<td id="menuCC" class="<%=strClass%>" <%=strJSEvents%>>
								<%=strALinkStart%>
									<img src="/xtra/images/icon_copyCommission.gif" width="16" height="15" alt="" align="absmiddle">Kopier oppdrag
								<%=strALinkEnd%>
							</td>
							
							<% 
							If  ((HasUserRight(ACCESS_TASK, RIGHT_WRITE) = true) and (AToolbarAtrib(3,0) = "1") and clng(lngOppdragID) > 0 ) Then 
								strALinkStart = AToolbarAtrib(3,1)
								strALinkEnd = AToolbarAtrib(3,2)							
								strClass = "menu"
								strJSEvents = "onMouseOver='menuOver(this.id);' onMouseOut='menuOut(this.id);'"							
							else
								strALinkStart = ""
								strALinkEnd = ""
								strClass = "menu disabled"
								strJSEvents = ""
							end if 
							%>
							<td id="menu4" class="<%=strClass%>" <%=strJSEvents%>>
								<%=strALinkStart%>
								<img src="/xtra/images/icon_AddToConsultant.gif" width="18" height="15" alt="" align="absmiddle">Tilknytt
								<%=strALinkEnd%>
							</td>
							<% 
							If  ((AToolbarAtrib(4,0) = "1") and clng(lngOppdragID) > 0 ) Then 
								strALinkStart = AToolbarAtrib(4,1)
								strALinkEnd = AToolbarAtrib(4,2)							
								strClass = "menu"
								strJSEvents = "onMouseOver='menuOver(this.id);' onMouseOut='menuOut(this.id);'"							
							else
								strALinkStart = ""
								strALinkEnd = ""
								strClass = "menu disabled"
								strJSEvents = ""
							end if 
							%>
							<td id="menu5" class="<%=strClass%>" <%=strJSEvents%>>
								<%=strALinkStart%>
								<img src="/xtra/images/icon_activities.gif" alt="" width="18" height="15" border="0" align="absmiddle">Aktiviteter
								<%=strALinkEnd%>
							</td>
						 
						</tr>
					</table>
				</td>
				<td class="right">
				<!--#include file="contentToolsMenu.asp"-->
				</td>
			</tr>
		</table>
	</div>