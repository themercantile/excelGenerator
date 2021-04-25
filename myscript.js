/* This is the processing part of the Create Excel Formulas project.
We're going to use jQuery as well to make things easier.
Need to provide instructions in the first div section and then get them to input the desired stuff in the second div section. 
I'm using the following 2 lines as guides on the correct formats
https://zoomies.pupper.org/MYRECOD122/displayme.jizz
start chrome "https://zoomies.pupper.org/", A1, "/displayme.jizz"
*/

function createURL() {
  let fullURL = document.getElementById("fullURL").value;
  console.log(fullURL);
  let recordVar = document.getElementById("recordVar").value;
  if (fullURL.search(recordVar)===-1) { // This line is searching for the string in the URL. If it isn't there a -1 is returned.
    document.getElementById("outURL").innerHTML = "ERROR!";
    document.getElementById("batchURL").innerHTML = "ERROR!";
    return;
  }
  console.log(recordVar);
  let excelCell = "\", " + document.getElementById("excelCell").value + ", \"";
  console.log(excelCell);
  let newURL = "=HYPERLINK(CONCATENATE(\"" + fullURL.replace(recordVar, excelCell) + "\"), \"Open Hyperlink\")";
  console.log(newURL);
  let chromeURL = "=CONCATENATE(\"start chrome " + fullURL.replace(recordVar, excelCell) + "\")";
  document.getElementById("outURL").innerHTML = newURL;
  document.getElementById("batchURL").innerHTML = chromeURL;
}

function createSkype() {
  let skypeContact = document.getElementById("excelEmailCell").value;
  let skypeURL = "=HYPERLINK(CONCATENATE(\"skype:\"," + skypeContact + ",\"?chat\"), \"Initiate Skype Chat\")";
  document.getElementById("skypeURL").innerHTML = skypeURL;
}
