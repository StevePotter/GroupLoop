<script language="JavaScript">
<!--
function getSelectValue( object ){
	var CurrentSelection, MasterMenuValue;
	CurrentSelection = object.selectedIndex;
	MasterMenuValue = object.options[CurrentSelection].value;

	return MasterMenuValue;
}


function MM_preloadImages() { //v2.0
  if (document.images) {
    var imgFiles = MM_preloadImages.arguments;
    if (document.preloadArray==null) document.preloadArray = new Array();
    var i = document.preloadArray.length;
    with (document) for (var j=0; j<imgFiles.length; j++) if (imgFiles[j].charAt(0)!="#"){
      preloadArray[i] = new Image;
      preloadArray[i++].src = imgFiles[j];
  } }
}


function DeleteBox( Msg, Deletelink )
{
	var where_to= confirm(Msg);
	if (where_to== true) {
		window.location=Deletelink;
	}
}

function openWindow(windowURL,windowName,windowWidth,windowHeight) {
	window.name = 'parentWnd';
	newWindow = window.open(windowURL,windowName,'width='+650+',toolbar=0,location=0,directories=0,status=0,menuBar=0,scrollBars=yes,resizable=1');
	newWindow.focus();
}

function Redirect( Link )
{
	window.location=Link
}

function show(object) {
    if (document.getElementById && document.getElementById(object) != null)
         node = document.getElementById(object).style.display='';
    else if (document.layers && document.layers[object] != null)
        document.layers[object].display = '';
    else if (document.all)
        document.all[object].style.display = '';
}

function hide(object) {
    if (document.getElementById && document.getElementById(object) != null)
         node = document.getElementById(object).style.display='none';
    else if (document.layers && document.layers[object] != null)
        document.layers[object].display = 'none';
    else if (document.all)
         document.all[object].style.display = 'none';
}

function isVisible(object) {
	var isVisible;
    if (document.getElementById && document.getElementById(object) != null)
         isVisible = document.getElementById(object).style.display=='';
    else if (document.layers && document.layers[object] != null)
        isVisible = document.layers[object].display == '';
    else if (document.all)
         isVisible = document.all[object].style.display == '';
	return isVisible;
}


function switchdisplay(object){
	if (isVisible(object))
		hide(object);
	else
		show(object);
}


function changeWord(wordID, newVal) {
    allElements = document.all;
    for (i=0; i<allElements.length; i++) {
		if (allElements(i).id==wordID)
			allElements(i).innerText=newVal;
    }
}
//-->

</SCRIPT>
