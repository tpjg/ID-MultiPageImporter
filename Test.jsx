// Test.jsx
// Collection of snippets testing InDesign javascript things

#target indesign;

var appVersion = parseInt(app.version);
// Only works in CS3+
if(appVersion >= 5)
{
	var oldInteractionPref = app.scriptPreferences.userInteractionLevel;
	app.scriptPreferences.userInteractionLevel = UserInteractionLevels.interactWithAll;
}
else
{
	alert("Features used in this script will only work in InDesign CS3 or later.");
	exit(-1);
}

// Set the next line to false to not use prefs
var usePrefs = true;

// Set default prefs
var prefOne = 0;
var prefTwo = 1;


// Look for and read prefs file
prefsFile = File((Folder(app.activeScript)).parent + "/testprefs.txt");
if(!prefsFile.exists)
{
	savePrefs(true);
}	
else
{
	readPrefs();
}


if(app.documents.length == 0)
{
    alert("No document open")
    exit(1)
}
var theDoc = app.activeDocument;
// Save zero point for later restoration
var oldZero = theDoc.zeroPoint;
// set the zero point to the origin
theDoc.zeroPoint = [0,0];
// Save ruler origin for later restoration
var oldRulerOrigin = theDoc.viewPreferences.rulerOrigin;
// set the ruler origin to page or all PDFs will be placed on first page of spreads
theDoc.viewPreferences.rulerOrigin = RulerOrigin.pageOrigin;

// Check retrieving path of open Document
alert(theDoc.fullName.name);
alert(theDoc.filePath.fsName);

// Check connecting to a socket and reaading something
var tcp = new Socket;
ok = tcp.open("127.0.0.1:2099")
tcp.writeln("DEST: "+theDoc.filePath.fsName)
recv = tcp.readln()
alert(recv + "(" + tcp.error + ")")
//tcp.close()

function progress(steps) {
    var b, t, w;

    w = new Window("palette", "Progress", undefined, {closeButton: false});
    t = w.add("statictext");
    t.preferredSize = [450, -1]; // 450 pixels wide, default height.

    if (steps) {
        b = w.add("progressbar", undefined, 0, steps);
        b.preferredSize = [450, -1]; // 450 pixels wide, default height.
    }

    progress.close = function () {
        w.close();
    };

    progress.increment = function () {
        b.value++;
    };

    progress.message = function (message) {
        t.text = message;

    };

    w.show();
}

// Start experiments...
if (false) {
    len = theDoc.hyperlinkTextSources.length;
    progress(5);
    //alert("test: " + len + "textsources");
    for (var i=0;i<len;i++){
        ts = theDoc.hyperlinkTextSources[i];
        progress.message("ts: " + ts.sourceText.contents + " < " + ts.parent + " ( " + ts.name);
        progress.increment();
        //alert("ts: " + ts.sourceText.contents + " < " + ts.parent + " ( " + ts.name);
        $.sleep(1000);
        //break;
        if (i==3) {
            alert("exiting")
            exit();
        }
    }
}

len = theDoc.hyperlinks.length
progress(len)
for (var i=0;i<len;i++){ 
    link = theDoc.hyperlinks[i];
    progress.message("link: " + link.source.sourceText.contents);
    progress.increment();
    if (link.destination instanceof HyperlinkURLDestination) {
        progress.message("URL: " + link.destination.destinationURL)
        tcp.writeln("FETCH: " + link.destination.destinationURL)
        recv = tcp.readln()
        if (recv.substr(0,6) === "FILE: "){
            progress.message(recv)
        } else {
            alert(recv + "(" + tcp.error + ")")
        }
        //exit();
    } else {
        //alert(link.name + " + " + link.label + " + " + link.destination);
    }
    $.sleep(200);
    if (i==3) {
        exit();
    }
}





// Save prefs and then restore original app/doc settings
savePrefs(false);
restoreDefaults(true);

// THE END OF EXECUTION
exit();

// function to restore saved settings back to originals before script ran
// extras parameter is for exiting at different areas of script:
// false: prior to doing anything
// true: end of script or reading PDF file size
function restoreDefaults(extras)
{
	app.scriptPreferences.userInteractionLevel = oldInteractionPref;
	if(extras == true)
	{
		theDoc.zeroPoint = oldZero;
		theDoc.viewPreferences.rulerOrigin = oldRulerOrigin;
	}
}

// function to read prefs from a file
function readPrefs()
{
	if(usePrefs)
	{
		try
		{
			prefsFile.open("r");
            prefOne = Number(prefsFile.readln() );
            prefTwo = Number(prefsFile.readln() );
			prefsFile.close();
		}
		catch(e)
		{
			throwError("Could not read preferences: " + e, false, 2, prefsFile);
		}
	}
}

// function to save prefs to a file
function savePrefs(firstRun)
{
	if(usePrefs)
	{
		try
		{
			var newPrefs = prefOne + "\n" +
            prefTwo + "\n"
			prefsFile.open("w");
			prefsFile.write(newPrefs);
			prefsFile.close();
		 }
		catch(e)
		{
			throwError("Could not save preferences: " + e, false, 2, prefsFile);
		}
	}
}

// Error function
function throwError(msg, pdfError, idNum, fileToClose)
{	

	if(fileToClose != null)
	{
		fileToClose.close();
	}
	
	if(pdfError)
	{
		// Throw err to be able to turn page numbering off
		throw Error("dummy");
	}
	else
	{
		alert("ERROR: " + msg + " (" + idNum + ")", "MultiPageImporter Script Error");
		exit(idNum);
	}
}
