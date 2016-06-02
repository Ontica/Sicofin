function writeDate(itemName, useSpanish, useMilitar) {
  var aMonthsSpanish = new Array("enero","febrero","marzo","abril","mayo","junio","julio","agosto","septiembre","octubre","noviembre","diciembre");
  var aMonthsEnglish = new Array("January","February","March","April","May","June","July","August","September","October","November","December");
  var aDaysSpanish   = new Array("Domingo","Lunes","Martes","Miércoles","Jueves","Viernes","Sábado");
  var aDaysEnglish   = new Array("Sunday","Monday","Tuesday","Wednesday","Thursday","Friday","Saturday");
  var today          = new Date();

  if (useSpanish) {
    //document.all[itemName].innerText = aDaysSpanish[today.getDay()] + " " + today.getDate() + " de " + aMonthsSpanish[today.getMonth()] + " del " + today.getFullYear() + getClock(useMilitar);
    document.all[itemName].innerText = aDaysSpanish[today.getDay()] + " " + today.getDate() + " de " + aMonthsSpanish[today.getMonth()] + " / " + getClock(useMilitar);
  } else {
    //document.all[itemName].innerText = aDaysEnglish[today.getDay()] + ", " + aMonthsEnglish[today.getMonth()] + " " + today.getDate() + ", " + today.getFullYear() + getClock(useMilitar);
    document.all[itemName].innerText = aDaysEnglish[today.getDay()] + ", " + aMonthsEnglish[today.getMonth()] + " " + today.getDate() + " / " + getClock(useMilitar);
  }
  window.setTimeout("writeDate('" + itemName + "', " + useSpanish + ", " + useMilitar + ");", 999);
}

function getClock(useMilitar) {
  var hours, minutes, ap;
  var intHours, intMinutes;
  var today;
  today      = new Date();
  intHours   = today.getHours();
  intMinutes = today.getMinutes();

  if (!useMilitar) { 
      if (intHours == 0) {
       hours = "12";
       ap = "a.m.";
    } else if (intHours < 12) { 
       hours = intHours;
       ap = "a.m.";
    } else if (intHours == 12) {
       hours = "12";
       ap = "p.m.";
    } else {
       intHours = intHours - 12;
       hours = intHours;
       ap = "p.m.";
    }
    minutes = intMinutes;
  } else {
     hours = intHours;
     minutes = intMinutes;
     ap = "hrs.";
  }

  if (intHours < 10 && intHours != 0) {
    hours = "0" + hours;
  }

  if (intMinutes < 10) {
    minutes = "0" + minutes;
  }    
  return (hours + ":" + minutes + ' ' + ap);
}



function launchToDesktop() {

  window.name = "newshome";

  if (((parseInt(navigator.appVersion)) >= 3) || (navigator.userAgent.indexOf("MSIE 4.0")) >= 0) {
    linkwindow("/JAVA/popoff.html", "", 141, 309);
  } else  { 
    if (navigator.appName.indexOf("Netscape") >= 0) {
       linkwindow("/JAVA/popoff.html", "", 141, 309);
    } else {
      linkwindow("/JAVA/popoff.html", "", 128, 273);
    }
  }
}


function randomBanner() {
  var randomnumber = Math.random();

  if (browserVer == 1) {
    i = Math.round( (i - 1) * randomnumber) + 1;
    document.banner.src = eval("banner" + i + ".src");
  }
}
