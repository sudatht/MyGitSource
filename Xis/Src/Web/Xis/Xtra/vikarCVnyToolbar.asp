							<table cellpadding="0" cellspacing="2">
								<tr>
									<td class="menu" id="menu5" onMouseOver="menuOver(this.id);" onMouseOut="menuOut(this.id);"><img src="/xtra/images/icon_consultant.gif" width="18" height="15" alt="" align="absmiddle"><a href="/xtra/VikarVis.asp?vikarid=<%=lngVikarID%>" title="Vis vikar">Vis vikar</a></td>
									<td class="menu" id="menu6" onMouseOver="menuOver(this.id);" onMouseOut="menuOut(this.id);"><strong>CV</strong>&nbsp;<select id="cboCVChoice" onChange="javascript:Vis_CV(<%=lngVikarID%>);"><option value="0"></option><option value="1">se</option><option value="3">presentere</option></select></td>
								</tr>
							</table>