							<table cellpadding="0" cellspacing="2">
								<tr>
									<%
								if (blnShowHotList) then	
										if (len(strAddToHotlistLink)>0) and (ccAction = false) then									
										%>
										<td class="menu" id="menu20" onMouseOver="menuOver(this.id);" onMouseOut="menuOut(this.id);"><a href="<%=strAddToHotlistLink%>" title="Legg <%=strHotlistType%> til i Snarveier på Min Side">Legg til i Snarveier</td>
										<%
										else
										%>
										<td class="menu disabled" id="menu21" <a href="" title="Legg <%=strHotlistType%> til i Snarveier på Min Side">Legg til i Snarveier</td>
										<%										
										end if
									end if
									%>										
									<td class="menu right" id="menu22" onMouseOver="menuOver(this.id);" onMouseOut="menuOut(this.id);"><a href="javascript:window.print();" title="Skriv ut"><img src="/xtra/images/icon_print.gif" width="18" height="15" alt="" align="absmiddle">Skriv ut</a></td>
									<!--td nowrap class="right">
										&nbsp; Tekst-størrelse:
										<a href="#" onMouseOver="imgOn('img1');" onMouseOut="imgOff('img1');" onClick="fnIncreaseFontSize();" title="Større skrift"><img name="img1" src="/xtra/images/icon_fntLarge.gif" width="18" height="15" alt="Større skrift" align="absmiddle"></a>
										<a href="#" onMouseOver="imgOn('img2');" onMouseOut="imgOff('img2');" onClick="fnDecreaseFontSize();" title="Mindre skrift"><img name="img2" src="/xtra/images/icon_fntSmall.gif" width="18" height="15" alt="Mindre skrift" align="absmiddle"></a>
										<span id="sizeNow" style="display:none;"></span>
									</td-->
								</tr>
							</table>