<SCRIPT LANGUAGE="JavaScript">
<!--
var bV=parseInt(navigator.appVersion);
NS4=(document.layers) ? true : false;
var agt = navigator.userAgent.toLowerCase();
NS5 = (navigator.appVersion.charAt(0) >= "5") ? true : false;
IE4=(!NS5 && (document.all) && (bV>=4))?true:false;

ver4 = (NS4 || IE4 || NS5) ? true : false;
function navOver(imName,dir,over,bot) {
if (!document.images) {return};
whichIm = (NS4 && bot) ? document.elFooter.document.images[imName] : document.images[imName];
if (over) {whichIm.src = "../art/b_" + dir + "_over.gif"}
else {whichIm.src = "../art/b_" + dir + ".gif"}
}
function navPress(imName,dir,press,bot) {
whichIm = (NS4 && bot) ? document.elFooter.document.images[imName] : document.images[imName];
if (press) {whichIm.src = "../art/b_" + dir + "_down.gif"}
else {whichIm.src = "../art/b_" + dir + ".gif"}
}
function expandIt(){return}
function expandAll(){return}


function justDoIt() {
   suggestion = document.hq.suggestions.options[document.hq.suggestions.selectedIndex].text;
   if (suggestion != "Recommendations") {
      document.hq.q.value=suggestion;
      document.hq.submit();
   }
}

isExpanded = false;
isDynamic = true;
preExpand = '';
preExpand2 = '';
function getIndex(el) {
	ind = null;
	for (i=0; i<document.layers.length; i++) {
		whichEl = document.layers[i];
		if (whichEl.id == el) {
			ind = i;
			break;
		}
	}
	return ind;
}

function arrange() {
	nextY = document.layers[firstInd].pageY + document.layers[firstInd].document.height;
	for (i=firstInd+1; i<document.layers.length; i++) {
		whichEl = document.layers[i];
		if (whichEl.visibility != "hide") {
			whichEl.pageY = nextY;
			nextY += whichEl.document.height;
		}
	}
}

function initIt(){
  if (isDynamic) {
  	if (NS4) {
		for (i=0; i<document.layers.length; i++) {
			whichEl = document.layers[i];
			if (whichEl.id.indexOf("Child") != -1) whichEl.visibility = "hide";
		}
		arrange();
	} else if (NS5) {
	        tempColl = document.getElementsByTagName("div");
		for (i=0; i<tempColl.length; i++) {
			if (tempColl.item(i).className == "child") tempColl.item(i).style.display = "none";
		}		
	} else {
		tempColl = document.all.tags("div");
		for (i=0; i<tempColl.length; i++) {
			if (tempColl(i).className == "child") tempColl(i).style.display = "none";
		}
	}
        if (preExpand != '') {
	  expandIt(preExpand);
        }
	if (preExpand2 != '') {
	  expandIt(preExpand2);
        }
  }
}

function expandIt(el) {
	if (!ver4) return;
	if (IE4) {expandIE(el)} else if (NS5) {expandNS5(el)} else  {expandNS(el)}
}

function expandIE(el) { 
	whichEl = eval(el + "Child");
	whichIm = eval(el + "Im");
	if (whichEl.style.display == "none") {
		whichEl.style.display = "block";
		whichIm.src = "/images/music/triUp.gif";
	}
	else {
		whichEl.style.display = "none";
		whichIm.src = "/images/music/triDown.gif";
	}
}

function expandNS5(el) { 
	whichEl = document.getElementById(el + "Child");
	whichIm = document.getElementById(el + "Im");
	if (whichEl.style.display == "none") {
		whichEl.style.display = "block";
		whichIm.src = "/images/music/triUp.gif";		
	}
	else {
		whichEl.style.display = "none";
		whichIm.src = "/images/music/triDown.gif";
	}
}

function expandNS(el) {
	whichEl = eval("document." + el + "Child");
	whichIm = eval("document." + el + "Parent.document.images['imEx']");
	if (whichEl.visibility == "hide") {
		whichEl.visibility = "show";
		whichIm.src = "/images/music/triUp.gif";
	}
	else {
		whichEl.visibility = "hide";
		whichIm.src = "/images/music/triDown.gif";
	}
	arrange();
}

function showAll() {
	for (i=firstInd; i<document.layers.length; i++) {
		whichEl = document.layers[i];
		whichEl.visibility = "show";
	}
}

function showTop() {
       for (i=firstInd; i<document.layers.length; i++) {
	 whichEl = document.layers[i];
       }
}

with (document) {
	write("<STYLE TYPE='text/css'>");
	if (NS4) {
		write(".parent {position:absolute; visibility:show}");
		write(".child {position:absolute; visibility:hidden}");
		write(".regular {position:absolute; visibility:show}")
	}
	else {
		write(".child {display:none}")
	}
	write("</STYLE>");
}

window.onload = initIt;

isDynamic = true;
if (NS4) {
	firstEl = "soundParent";
	firstInd = getIndex(firstEl);
	showTop();
	arrange();
} 
//-->
</SCRIPT>