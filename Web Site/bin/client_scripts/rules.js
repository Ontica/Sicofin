/**********************************************/
/*   Operaciones con ventanas desplegables    */
/**********************************************/

var oCurrentFullScrollMenu = null;

function showOptionsWindow(oWindow) {
	if (typeof(oWindow.tacked) == 'undefined') {
		oWindow.tacked = false;
	}
	if ((oCurrentFullScrollMenu != null) && (oWindow != oCurrentFullScrollMenu)) {
		if (oCurrentFullScrollMenu.tacked == false) {
			oCurrentFullScrollMenu.style.display = 'none';
		}
	}
	if (oWindow.tacked == false) {
		if (oWindow.style.display == 'inline') {
			oWindow.style.display = 'none';
			oWindow.tacked = false;
			oCurrentFullScrollMenu = null;
		}	else {
			oWindow.style.display = 'inline';
			oWindow.tacked = false;
			oCurrentFullScrollMenu = oWindow;
		}
	}
	return false;
}

function showTackedWindows(aTackedWindows) {
	var i, j;
	
	for (i = 0; i < aTackedWindows.length; i++) {
		showOptionsWindow(document.all[aTackedWindows[i]]);
		tackOptionsWindow(document.all[aTackedWindows[i]]);
	}
}

function tackOptionsWindow(oWindow) {
	var oTackThumb, i;
  var aChildImages;
  var oTackedWindows = document.all.txtTackedWindows;
  if (arguments[1] != null) {
      oTackThumb = arguments[1];
  } else {
		aChildImages = oWindow.all.tags('IMG');
		for (i = 0; i < aChildImages.length ; i++) {
			if (aChildImages[i].src.indexOf('tack_') >= 0) {
				oTackThumb = aChildImages[i];
			}
		}
  } 
	if (oWindow.tacked == true) {
		oTackThumb.src = '/empiria/images/tack_white.gif';
		oTackThumb.alt = 'Fijar la ventana';
		oWindow.tacked = false;
		oTackedWindows.value = oTackedWindows.value.replace("'" + oWindow.id + "',", "")
		oTackedWindows.value = oTackedWindows.value.replace(",'" + oWindow.id + "'", "")
		oTackedWindows.value = oTackedWindows.value.replace("'" + oWindow.id + "'", "")
	} else {
		oTackThumb.src = '/empiria/images/tack_down_white.gif';
		oTackThumb.alt = 'Desclavar la ventana';
		oWindow.tacked = true;
		if (oTackedWindows.value == '') {
			oTackedWindows.value = "'" + oWindow.id + "'";
		} else {
			oTackedWindows.value += ",'" + oWindow.id + "'";
		}
	}
}

function closeOptionsWindow(oWindow, oTackThumb) {
	if (oWindow.tacked) {
		tackOptionsWindow(oWindow, oTackThumb);
	}
	showOptionsWindow(oWindow);
}

function notAvailable() {
	alert("Por el momento esta opción no está disponible");
	return false;
}


function cmdCheckSpelling_onclick() {
	alert("Por el momento la revisión ortográfica no está disponible.\n\nGracias.");
}

function showCalendar(oSource) {		
	alert("El calendario no está disponible en este momento.\n\nGracias.");	
	return false;
}

/***********************************/
/*   Operaciones con checkboxes    */
/***********************************/

function countCheckBoxes(sCheckBoxName) {
	var i = 0, counter = 0;
	
	if (typeof(document.all[sCheckBoxName]) == 'undefined') {
		return 0;
	}	
	if (document.all[sCheckBoxName].length != null) {
		for (i = 0 ; i < document.all[sCheckBoxName].length ; i++) {
			if (document.all[sCheckBoxName](i).checked) {
				counter++;
			}
		}
	} else {
		if (document.all[sCheckBoxName].checked) {
			counter++;
		}
	}
	return counter;
}

function selectCheckBoxes(sCheckBoxName, bCheck) {
	var i = 0;
	
	if (document.all[sCheckBoxName].length != null) {
		for (i = 0 ; i < document.all[sCheckBoxName].length ; i++) {
			document.all[sCheckBoxName](i).checked = bCheck;
			if (bCheck) {
				document.all[sCheckBoxName](i).parentElement.parentElement.className = 'applicationTableSelectedRow';
			} else {	
				document.all[sCheckBoxName](i).parentElement.parentElement.className = '';
			} 
		}		
	} else {
		document.all[sCheckBoxName].checked = bCheck;	
		if (bCheck) {
			document.all[sCheckBoxName].parentElement.parentElement.className = 'applicationTableSelectedRow';
		} else {	
			document.all[sCheckBoxName].parentElement.parentElement.className = '';
		} 			
	}
	return true;	
}

/****************************************/
/*   Operaciones con tablas arboreas    */
/****************************************/

function collapseAll() {
	var oObject = event.srcElement;
	var oTableSections, oTemp, oRows;
	var nLength, i, j;	
	
	oTableSections = getObjectParent(oObject, 'TABLE').tBodies;
	nLength = oTableSections.length;	
	for (i = 0; i < nLength; i++) {
		oTemp = oTableSections[i];
		oTemp.value = Math.abs(oTemp.value);
		oRows = oTemp.rows;
		if ((oRows.length > 0) && (oRows[0].cells[0].all.tags('IMG').length > 0)) {
			oTemp.all.tags('IMG')[0].src = '/empiria/images/expanded.gif';
		}
		if (oTemp.value > 1) {
			oTemp.style.display = 'none';
		}
	}
	return false;
}

function expandAll() {
	var oObject = event.srcElement;
	var oTableSections, oTemp, oRows;
	var nLength, i, j;	
	
	oTableSections = getObjectParent(oObject, 'TABLE').tBodies;
	nLength = oTableSections.length;
	for (i = 0; i < nLength; i++) {		
		oTemp = oTableSections[i];
		oTemp.style.display = 'inline';
		oTemp.value = Math.abs(oTemp.value);
		oRows = oTemp.rows;
		if ((oRows.length > 0) && oRows[0].cells[0].all.tags('IMG').length > 0) {
			if (oTemp.id == '') {
				oRows[0].cells[0].all.tags('IMG')[0].src = '/empiria/images/collapsed.gif';
			} else {
				oRows[0].cells[0].all.tags('IMG')[0].src = '/empiria/images/expanded.gif';
			}			
		}
	}
	return false;
}

function outline() {
	var oObject = event.srcElement;
	if (oObject.src.indexOf('expanded.gif') >= 0) {
		oObject.src = '/empiria/images/collapsed.gif';
		oObject.alt = 'Contraer todo';
		expandAll();	
	} else {
		oObject.src = '/empiria/images/expanded.gif';
		oObject.alt = 'Expandir todo';
		collapseAll();
	}
}

function outliner() {
	var oObject = event.srcElement;
	var bExpand, nSectionLevel, nStartSection, i, j, nObjectId;
	var oTableSections, oStartSection;
	
	if (oObject.src.indexOf('expanded.gif') >= 0) {
		bExpand = true;
		oObject.src = '/empiria/images/collapsed.gif';
	} else  {
		bExpand = false;
		oObject.src = '/empiria/images/expanded.gif';
	}
	
	oStartSection  = getObjectParent(oObject, 'TBODY');
	oTableSections = getObjectParent(oStartSection, 'TABLE').tBodies;
	nSectionLevel  = Number(oStartSection.value);
	nStartSection  = sectionIndex(oTableSections, oStartSection);
  
	if (bExpand) {
	  if (oTableSections[nStartSection].id != '') {
			insertSectionInfo(oTableSections[nStartSection + 1], oTableSections[nStartSection].id);
			oTableSections[nStartSection].id = '';
	  }
		for (i = nStartSection + 1; i < oTableSections.length; i++) {
			if ((nSectionLevel + 1) == Math.abs(oTableSections[i].value)) {
				oTableSections[i].value = Math.abs(oTableSections[i].value);
				oTableSections[i].style.display = 'inline';
				continue;
			}
			if ((nSectionLevel + 1) < Math.abs(oTableSections[i].value)) {
				if ((oTableSections[i].value < 0) && (parentSection(oTableSections, oTableSections[i], i).style.display == 'inline')) {
				  oTableSections[i].value = Math.abs(oTableSections[i].value);
					oTableSections[i].style.display = 'inline';
				}
				continue;
			}
			break;
		}
	} else {
		for (i = nStartSection + 1; i < oTableSections.length; i++) {
			if ((nSectionLevel + 1) == Math.abs(oTableSections[i].value)) {
				oTableSections[i].value = Math.abs(oTableSections[i].value);
				oTableSections[i].style.display = 'none';
				continue;
			}
			if ((nSectionLevel + 1) < Math.abs(oTableSections[i].value)) {			  
				if (oTableSections[i].style.display == 'inline') {
					oTableSections[i].value = (-1 * Math.abs(oTableSections[i].value));
				}						
				oTableSections[i].style.display = 'none';				
				continue;
			}
			break;
		}
	}
	return false;
}

function parentSection(oSections, oSection, nSectionIndex){
	var nLevel, i;
		
	nLevel =  Math.abs(oSection.value) - 1;	 
	for (i = nSectionIndex - 1; i > 0; i--) {	
		if (Math.abs(oSections[i].value) == nLevel) {
			return (oSections[i]);
		}
	}
	return (oSection);
}

function sectionIndex(oBodiesCollection, oSection) {
	var i;
	
	for (i = 0; i < oBodiesCollection.length; i++) {
		if (oBodiesCollection[i] == oSection) {
			return i;
		}
	}
	return (-1);
}

function selectItems(oSource) {
	var oStartSection = getObjectParent(oSource, 'TBODY');
	var oTableSections, oTableSection, nSectionLevel, nStartSection, i, j;

	oTableSections = getObjectParent(oStartSection, 'TABLE').tBodies;
	nSectionLevel  = oStartSection.value;
	nStartSection  = sectionIndex(oTableSections, oStartSection);
  
  oTableSection = oTableSections[nStartSection];
  oTableSection.rows[0].all.tags('IMG')[0].src = '/empiria/images/collapsed.gif';
	for (i = 1; i < oTableSection.rows.length; i++) {
		oTableSection.rows[i].style.display = 'inline';
		oTableSection.rows[i].className = 'applicationTableSelectedRow';
		oTableSection.rows[i].all.tags('INPUT')[0].checked = true;
	}
    
	for (i = nStartSection + 1; i < oTableSections.length; i++) {
		oTableSection = oTableSections[i];
		if (nSectionLevel < Math.abs(oTableSection.value)) {
			oTableSection.value = Math.abs(oTableSection.value);
			oTableSection.style.display = 'inline';
			oTableSection.rows[0].all.tags('IMG')[0].src = '/empiria/images/collapsed.gif';
			for (j = 1; j < oTableSection.rows.length; j++) {
				oTableSection.rows[j].style.display = 'inline';			
				oTableSection.rows[j].className = 'applicationTableSelectedRow';
				oTableSection.rows[j].all.tags('INPUT')[0].checked = true;
			}
		} else {
			break;
		}
	}
	return false;
}

function selectRow() {
	var oSource = event.srcElement;
	
	getObjectParent(oSource, 'TR').className = ((oSource.checked) ? 'applicationTableSelectedRow' : '');
}

function unSelectItems(oSource) {
	var oStartSection = getObjectParent(oSource, 'TBODY');
	var oTableSections, oTableSection, nSectionLevel, nStartSection, i, j;
	
	oTableSections = getObjectParent(oStartSection, 'TABLE').tBodies;
	nSectionLevel  = oStartSection.value;
	nStartSection  = sectionIndex(oTableSections, oStartSection);

  oTableSection = oTableSections[nStartSection];
	for (i = 1; i < oTableSection.rows.length; i++) {
		oTableSection.rows[i].className = '';
		oTableSection.rows[i].all.tags('INPUT')[0].checked = false;
	}
	
	for (i = nStartSection + 1; i < oTableSections.length; i++) {
		oTableSection = oTableSections[i];
		if (nSectionLevel < Math.abs(oTableSection.value)) {
			for (j = 1; j < oTableSection.rows.length; j++) {
				oTableSection.rows[j].className = '';
				oTableSection.rows[j].all.tags('INPUT')[0].checked = false;
			}
		} else {
			break;
		}
	}
	return false;
}

/****************************************/
/*   Operaciones generales							*/
/****************************************/

function createWindow(oWindow, sURL, sPars) {
	var oTempWindow;

	if (oWindow == null || oWindow.closed) {
		oTempWindow = window.open(sURL, '_blank', sPars);
		return oTempWindow;
	} else {
		oWindow.focus();
		oWindow.navigate(sURL);
		return oWindow;
	}
}

function getObjectParent(oObject, sParentTag) {
	var oTemp = oObject;
	
	if (sParentTag != '') {
		while (oTemp != null) {
			if (oTemp.tagName != sParentTag) {
				oTemp = oTemp.parentElement;
			} else {
				return (oTemp);
			}
		}
	} else {
		return (oObject.parentElement);
	}
}

function setCursor(sCursor) {
	document.body.style.cursor = sCursor;
}

function showHelp(sItem) {	
	var sOptions = 'height=240px,width=380px,status:false,titlebar:false,location:false,toolbar:false,resizable:0,status:0';
	var sURL = '../../help/inline_help.asp?item=';
	window.open(sURL + sItem, null, sOptions);
	return false;
}

function unloadWindows(oWindow) {
	var i;

	if (oWindow != null && !oWindow.closed) {
		oWindow.close();
	}	
	for (i = 1 ; i < arguments.length; i++) {
		if (arguments[i] != null && !arguments[i].closed) {
			arguments[i].close();
		}
	}
}