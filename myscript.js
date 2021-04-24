/* This is the processing part of the Create Excel Formulas project.
We're going to use jQuery as well to make things easier.
Need to provide instructions in the first div section and then get them to input the desired stuff in the second div section. 
headSection and inputSection are the IDs for the divs.
=HYPERLINK(CONCATENATE("https://zoomies.pupper.org/", A1, "/displayme.jizz"), "CLICK")
https://zoomies.pupper.org/MYRECOD122/displayme.jizz
start chrome "https://zoomies.pupper.org/", A1, "/displayme.jizz"
*/

function createURL() {
  let fullURL = document.getElementById("fullURL").value;
  console.log(fullURL);
  let recordVar = document.getElementById("recordVar").value;
  if (fullURL.search(recordVar)===-1) {
    document.getElementById("outURL").innerHTML = "ERROR!";
    return;
  }
  console.log(recordVar);
  let excelCell = "\", " + document.getElementById("excelCell").value + ", \"";
  console.log(excelCell);
  let newURL = "=HYPERLINK(CONCATENATE(\"" + fullURL.replace(recordVar, excelCell) + "\"), \"Open Hyperlink\")";
  console.log(newURL);
  let chromeURL = "=CONCATENATE(\"start chrome " + fullURL.replace(recordVar, excelCell) + ")\"";
  //start chrome https://zoomies.pupper.org/", A1, "/displayme.jizz
  
  //=CONCATENATE("start chrome https://zoomies.pupper.org/", A1, "/displayme.jizz"
  document.getElementById("outURL").innerHTML = newURL;
  document.getElementById("batchURL").innerHTML = chromeURL;
}
