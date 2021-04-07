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


// Stuff to keep track of appendices
var appendixPages = -1;
var currentAppendix = null;
var appendixNumber = 0;
// Keep track of latest added appendix textpoint
var returnTextPoint;

// Stuff that MultiPageImporter defines globally
var docStartPG = 1;
// Do not change anything after this line!
// removed 6/25/08: var indUpdateType = 0;
var cropType = 0;
var PDF_DOC = "PDF";
var IND_DOC = "InDesign";
var tempObjStyle = null;
//var getout;
var rotateValues = [0,90,180,270];
var positionValuesAll = ["Top left", "Top center", "Top right", "Center left",  "Center", "Center right", "Bottom left",  "Bottom center",  "Bottom right"];
var noPDFError = true;
// Set default prefs
var pdfCropType = 0;
var indCropType = 1;
var offsetX = 0;
var offsetY = 0;
var doTransparent = 1;
var placeOnLayer = 0;
var fitPage = 1;
var keepProp = 1;
var addBleed = 1;
var fitMargin = 1; // Place within margins instead of fitting to page
var paragraphStyleName = "Appendix header"; // Name of the style to apply to text preceding the placed pages
var ignoreErrors = 0;
var percX = 100;
var percY = 100;
var mapPages = 0;
var reverseOrder = 0;
var rotate = 0;
var positionType = 0; //4; // 4 = center





if(app.documents.length == 0)
{
    alert("No document open");
    exit();
}
theDoc = app.activeDocument;
// Save zero point for later restoration
var oldZero = theDoc.zeroPoint;
// set the zero point to the origin
theDoc.zeroPoint = [0,0];
// Save ruler origin for later restoration
var oldRulerOrigin = theDoc.viewPreferences.rulerOrigin;
// set the ruler origin to page or all PDFs will be placed on first page of spreads
theDoc.viewPreferences.rulerOrigin = RulerOrigin.pageOrigin;

// Check retrieving path of open Document
docName = theDoc.fullName.fsName;
docName = docName.substr(0,docName.lastIndexOf("."));
//alert(docName);
//alert(theDoc.fullName.fsName);
//alert(theDoc.filePath.fsName);
var nummm = 0;

var theBook =  app.activeBook;

// Check connecting to a socket and reaading something
var tcp = new Socket;
ok = tcp.open("127.0.0.1:2099")
tcp.writeln("DEST: "+theDoc.filePath.fsName)
recv = tcp.readln()
if (recv !== "OK"){
    alert(recv + "(" + tcp.error + ")")
}
//tcp.close()

function progress(steps) {
    var b, t, w;

    w = new Window("palette", "Progress", undefined, {closeButton: false});
    t = w.add("statictext");
    t.preferredSize = [600, -1]; // default height.
    tsub1 = w.add("statictext");
    tsub1.preferredSize = [600, -1];
    tsub2 = w.add("statictext");
    tsub2.preferredSize = [600, -1];

    if (steps) {
        b = w.add("progressbar", undefined, 0, steps);
        b.preferredSize = [600, -1]; // default height.
    }

    progress.close = function () {
        w.close();
    };

    progress.increment = function () {
        b.value++;
    };

    progress.message = function (message) {
        t.text = message;
        tsub1.text = "";
        tsub2.text = "";
    };

    progress.msgSub1 = function (message) {
        tsub1.text = message;
    };

    progress.msgSub2 = function (message) {
        tsub2.text = message;
    };

    w.show();
}

// Start experiments...

addFileToAppendix();

len = theDoc.hyperlinks.length
progress(len)
for (mi=0;mi<len;mi++){ 
    link = theDoc.hyperlinks[mi];
    progress.message("Link: " + link.source.sourceText.contents);
    progress.increment();
    if (link.destination instanceof HyperlinkURLDestination) {
        progress.msgSub1("URL: " + link.destination.destinationURL)
        tcp.writeln("FETCH: " + link.destination.destinationURL)
        recv = tcp.readln()
        if (recv.substr(0,6) === "FILE: "){
            progress.msgSub2(recv)

            existingDest = getExistingHyperlinkTextDestination(theDoc,recv);
            if ((existingDest != null)){ //} && (existingDest.name === recv)){
                progress.message(recv + "already exists");
                setLinkProperties(link, existingDest);
                continue;
            }
            openFileAndPlaceIt(recv)
            if ((true) && (returnTextPoint)) {
                newDestination = theDoc.hyperlinkTextDestinations.add(returnTextPoint);
                newDestination.name = recv;
                setLinkProperties(link, newDestination);
                returnTextPoint = null;
            }
        } else {
            alert(recv + "(" + tcp.error + ")")
        }
    }
}

function getExistingHyperlinkTextDestination(theDoc, recv){
    for (j=0; j<theDoc.hyperlinks.length; j++){
        if (theDoc.hyperlinks[j].destination instanceof HyperlinkTextDestination){
            if (theDoc.hyperlinks[j].destination.name === recv){
                return theDoc.hyperlinks[j].destination;
            }
        }
    }
    return null;
}

function setLinkProperties(link, newDestination){
    link.destination = newDestination;
    link.borderColor = UIColors.LIGHT_BLUE;
    link.borderStyle = HyperlinkAppearanceStyle.SOLID;
    link.width = HyperlinkAppearanceWidth.MEDIUM;
    link.highlight = HyperlinkAppearanceHighlight.OUTLINE;
    link.visible = true;
}


var theFileToBePlaced;

function openFileAndPlaceIt(recv){
    fileNameToBePlaced = recv.substr(6);
    theFileToBePlaced = new File(fileNameToBePlaced);
    progress.msgSub2("Placing " + File.decode(theFileToBePlaced.name));

    // In "production" single UNDO, in Debug just run it ...
    if (false){
        app.doScript(internalPlaceIt, ScriptLanguage.JAVASCRIPT, undefined, UndoModes.ENTIRE_SCRIPT, "Place " + fileNameToBePlaced);
    } else {
        internalPlaceIt();
    }
    theFileToBePlaced.close();

    // Split appendices into documents of at most a certain number of pages
    if (currentAppendix.pages.length < 50){
        currentAppendix.pages.add(LocationOptions.AT_END);
        docStartPG = currentAppendix.pages.length;
    } else {
        currentAppendix.save();
        currentAppendix = null;
        addFileToAppendix();
        docStartPG = 1;
    }
    //currentAppendix.windows.add(); // Unhide hidden window
}

function internalPlaceIt(){
    placePDFat(theFileToBePlaced);
}

function addFileToAppendix(){
    if (currentAppendix == null){
        currentAppendix = app.open(File("/Users/timo/Documents/empty-memo.indt"),true); //false); //TODO: choose to hide or not
        appendixName = docName + "-app" + appendixNumber.toString() + ".indd";
        appendixNumber++;
        f = new File(appendixName);
        currentAppendix.save(f);
        theBook.bookContents.add(f)
    }
}

// Save prefs and then restore original app/doc settings
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






// Copy taken from (modified) ID-MultiPageImporter.
// Modifications include placing PDF within the margins of a page, creating sections and links.
function placePDFat(theFile){

// Check  if cancel was clicked
if (theFile == null)
{
	// user clicked cancel, just leave
    return;
}
// Check if a file other than PDF or InDesign chosen
else if((theFile.name.toLowerCase().indexOf(".pdf") == -1 && theFile.name.toLowerCase().indexOf(".ind") == -1 && theFile.name.toLowerCase().indexOf(".ai") == -1 ))
{
    return;
	throwError("A PDF, PDF compatible AI or InDesign file must be chosen. Quitting...", false, 1, null);
}

var fileName = File.decode(theFile.name);

if((theFile.name.toLowerCase().indexOf(".pdf") != -1) || (theFile.name.toLowerCase().indexOf(".ai") != -1))
{
	// Premedia Systems/JJB Edit Start - 02/14/11 Modified PDFCrop constants to support ID CS3 through CS5 PDFCrop Types.
	if (appVersion > 6)
	{
		// CS5 or newer
		var cropTypes = [PDFCrop.cropPDF, PDFCrop.cropArt, PDFCrop.cropTrim, PDFCrop.cropBleed, PDFCrop.cropMedia, PDFCrop.cropContentAllLayers, PDFCrop.cropContentVisibleLayers];
		var cropStrings = ["Crop","Art","Trim","Bleed", "Media","All Layers Bounding Box","Visible Layers Bounding Box"];
	}
	else
	{
		// CS3 or CS4
		var cropTypes = [PDFCrop.cropContent, PDFCrop.cropArt, PDFCrop.cropPDF, PDFCrop.cropTrim, PDFCrop.cropBleed, PDFCrop.cropMedia];
		var cropStrings = ["Bounding Box","Art","Crop","Trim","Bleed", "Media"];
	}
	// Premedia Systems/JJB Edit End
	
	// Parse the PDF file and extract needed info
    try
	{
        var placementINFO = getPDFInfo(theFile, (app.documents.length == 0));
    }
	catch(e)
	{
		// Couldn't determine the PDF info, revert to just adding all the pages
		noPDFError = false;
		placementINFO = new Array();
		
		if(app.documents.length == 0)
		{
			var tmp = new Array();
			tmp["width"] = 612;
			tmp["height"] = 792;

			placementINFO["pgSize"]  = tmp;
		}
	}
	placementINFO["kind"] = PDF_DOC;
}
else
{
	var cropTypes = [ImportedPageCropOptions.CROP_CONTENT, ImportedPageCropOptions.CROP_BLEED, ImportedPageCropOptions.CROP_SLUG];
	var cropStrings = ["Page bounding box","Bleed bounding box","Slug bounding box"];
	// Get the InDesign doc's info
	var placementINFO = getINDinfo(theFile);
	placementINFO["kind"] = IND_DOC;
}

var currentLayer = currentAppendix.activeLayer;
var docPgCount = currentAppendix.pages.length;
var lastPageNumber = Number(currentAppendix.pages[docPgCount-1].name);

// Not using dialog box - just place everything
startPG = 1;
if (noPDFError) {
    endPG = placementINFO.pgCount;
} else {
    endPG = 9999;
}

// Add the new layer if requested
if(placeOnLayer)
{
	// Add random number to file name to be layer name.
	// Double check layer name doesn't exist and alter if it happens to be present for some reason
	var layerName = fileName + "_" + Math.round(Math.random() * 9999);
	var docLayers = currentAppendix.layers;
	for(i=0; i < docLayers.length; i++)
	{
		if (docLayers[i].name.indexOf(layerName) != -1 )
		{
			layerName += ("_" + Math.round(Math.random() * 9999));
		}
	}
	
	// Add the layer
	currentLayer = currentAppendix.layers.add({name:layerName});
}

// Save zero point for later restoration
var oldZero = currentAppendix.zeroPoint;
// set the zero point to the origin
currentAppendix.zeroPoint = [0,0];

// Save ruler origin for later restoration
var oldRulerOrigin = currentAppendix.viewPreferences.rulerOrigin;
// set the ruler origin to page or all PDFs will be placed on first page of spreads
currentAppendix.viewPreferences.rulerOrigin = RulerOrigin.pageOrigin;

// Get the Indy doc's height and width
var docWidth = currentAppendix.documentPreferences.pageWidth;
var docHeight = currentAppendix.documentPreferences.pageHeight;

// Set placement prefs
if(placementINFO.kind == PDF_DOC)
{
	with(app.pdfPlacePreferences)
	{
		transparentBackground = doTransparent;
		pdfCrop = cropTypes[cropType];
	}
}
else
{
	app.importedPageAttributes.importedPageCrop = cropTypes[cropType];
}

// Block errors if requested
if(ignoreErrors)
{
	app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;
}

// Create the Object Style to be applied to the placed pages.
tempObjStyle = currentAppendix.objectStyles.add();
styleName = "MultiPageImporter_Styler_" + Math.round(Math.random() * 9999);
for (ti=0; (currentAppendix.objectStyles.itemByName(styleName) != null) || ti>10; ti++){
    styleName = "MultiPageImporter_Styler_" + Math.round(Math.random() * 9999);
}
tempObjStyle.name =  styleName;
tempObjStyle.strokeWeight = 0; // Make sure there's no stroke
tempObjStyle.fillColor = "None"; // Make sure fill is none
tempObjStyle.enableAnchoredObjectOptions = true;

// Set the anchor properties
var tempAOS = tempObjStyle.anchoredObjectSettings;
tempAOS.anchoredPosition = AnchorPosition.ANCHORED;
tempAOS.spineRelative = false;
tempAOS.lockPosition = false;
tempAOS.verticalReferencePoint = AnchoredRelativeTo.PAGE_EDGE;
tempAOS.horizontalReferencePoint = AnchoredRelativeTo.PAGE_EDGE;
tempAOS.anchorXoffset = offsetX;
tempAOS.anchorYoffset = offsetY;

// Set the placement options based on user selected position
// The -1 is needed to get rectangle to move correctly when using the auto positioning of the object styles
// Could be a bug since just the left positions need the negative multiple (spine doesn't need the negative multiple)
switch(positionType)
{
	case 0: //  Top Left
		tempAOS.anchorXoffset *= -1;
		tempAOS.anchorPoint = AnchorPoint.TOP_LEFT_ANCHOR;
		tempAOS.verticalAlignment = VerticalAlignment.TOP_ALIGN;
		tempAOS.horizontalAlignment = HorizontalAlignment.LEFT_ALIGN;		
		break;
	case 1: // Top Center
		tempAOS.anchorPoint = AnchorPoint.TOP_CENTER_ANCHOR;
		tempAOS.verticalAlignment = VerticalAlignment.TOP_ALIGN;
		tempAOS.horizontalAlignment = HorizontalAlignment.CENTER_ALIGN;
		break;
	case 2: // Top Right
		tempAOS.anchorPoint = AnchorPoint.TOP_RIGHT_ANCHOR;
		tempAOS.verticalAlignment = VerticalAlignment.TOP_ALIGN;
		tempAOS.horizontalAlignment = HorizontalAlignment.RIGHT_ALIGN;
		break;
	case 3: // Middle Left
		tempAOS.anchorXoffset *= -1;
		tempAOS.anchorPoint = AnchorPoint.LEFT_CENTER_ANCHOR;
		tempAOS.verticalAlignment = VerticalAlignment.CENTER_ALIGN;
		tempAOS.horizontalAlignment = HorizontalAlignment.LEFT_ALIGN;
		break;
	case 4: // Center
		tempAOS.anchorPoint = AnchorPoint.CENTER_ANCHOR;
		tempAOS.verticalAlignment = VerticalAlignment.CENTER_ALIGN;
		tempAOS.horizontalAlignment = HorizontalAlignment.CENTER_ALIGN;
		break;
	case 5: // Middle Right
		tempAOS.anchorPoint = AnchorPoint.RIGHT_CENTER_ANCHOR;
		tempAOS.verticalAlignment = VerticalAlignment.CENTER_ALIGN;
		tempAOS.horizontalAlignment = HorizontalAlignment.RIGHT_ALIGN;
		break;
	case 6: // Bottom Left
		tempAOS.anchorXoffset *= -1;
		tempAOS.anchorPoint = AnchorPoint.BOTTOM_LEFT_ANCHOR;
		tempAOS.verticalAlignment = VerticalAlignment.BOTTOM_ALIGN;
		tempAOS.horizontalAlignment = HorizontalAlignment.LEFT_ALIGN;
		break;
	case 7: // Bottom Center
		tempAOS.anchorPoint = AnchorPoint.BOTTOM_CENTER_ANCHOR;
		tempAOS.verticalAlignment = VerticalAlignment.BOTTOM_ALIGN;
		tempAOS.horizontalAlignment = HorizontalAlignment.CENTER_ALIGN;
		break;
	case 8: // Bottom Right
		tempAOS.anchorPoint = AnchorPoint.BOTTOM_RIGHT_ANCHOR;
		tempAOS.verticalAlignment = VerticalAlignment.BOTTOM_ALIGN;
		tempAOS.horizontalAlignment = HorizontalAlignment.RIGHT_ALIGN;
		break;
	// 9 == separator
	case 10: // Top Relative to Spine
		tempAOS.spineRelative = true;
		tempAOS.anchorXoffset *= -1;
		tempAOS.anchorPoint = AnchorPoint.TOP_RIGHT_ANCHOR;
		tempAOS.verticalAlignment = VerticalAlignment.TOP_ALIGN;
		tempAOS.horizontalAlignment = HorizontalAlignment.RIGHT_ALIGN;
		break;
	case 11: // Middle Relative to Spine
		tempAOS.spineRelative = true;
		tempAOS.anchorXoffset *= -1;
		tempAOS.anchorPoint = AnchorPoint.RIGHT_CENTER_ANCHOR;
		tempAOS.verticalAlignment = VerticalAlignment.CENTER_ALIGN;
		tempAOS.horizontalAlignment = HorizontalAlignment.RIGHT_ALIGN;			
		break;
	case 12: //  Bottom Relative to Spine
		tempAOS.spineRelative = true;
		tempAOS.anchorXoffset *= -1;
		tempAOS.anchorPoint = AnchorPoint.BOTTOM_RIGHT_ANCHOR;
		tempAOS.verticalAlignment = VerticalAlignment.BOTTOM_ALIGN;
		tempAOS.horizontalAlignment = HorizontalAlignment.RIGHT_ALIGN;
		break;
}

addPages(placementINFO, docWidth, docHeight, docStartPG, startPG, endPG, theFile);


// Kill the Object style
tempObjStyle.remove();
restoreAppendixDefaults(true);

// THE END OF EXECUTION
return;
}


// Place the requested pages in the document
function addPages(placementINFO, docWidth, docHeight, docStartPG, startPG, endPG, theFile)
{
	var currentPDFPg = 0;
	var firstTime = true;
	var addedAPage = false;
	var zeroBasedDocPgCnt = currentAppendix.pages.length - 1;

	for(i = docStartPG - 1, currentInputDocPg = startPG; currentInputDocPg <= endPG; currentInputDocPg++, i++)
	{
		// Since docs (especially if used in books) not necessarily have pages.index == pages.name,
		// use pages.itemByName(String(i+1)) instead of pages[i].

		if(placementINFO.kind == PDF_DOC)
		{
			// Set the app's PDF placement pref's page number property to the current PDF page number
			app.pdfPlacePreferences.pageNumber = currentInputDocPg;
		}
		else
		{
			// Set the app's Imported Page placement pref's page number property to the current IND page number
			app.importedPageAttributes.pageNumber = currentInputDocPg;
		}

		if(i > (currentAppendix.pages.length-1))
		{
			// Make sure we have a page to insert into
			currentAppendix.pages.add(LocationOptions.AT_END);
			addedAPage = true;
		}
	
		var margins = currentAppendix.pages[i].marginPreferences;

		var bounds = [margins.top, margins.left, docHeight-margins.bottom, docWidth-margins.right]; //[0,0,20,20];
		if (fitMargin){
			bounds = [margins.top, margins.left, docHeight-margins.bottom, docWidth-margins.right];
			// TG: add filename on first placement above the margin and start a new section
			if (firstTime){
				var myTextFrame = currentAppendix.pages[i].textFrames.add();
				// From documentation:
				// When document.documentPreferences.facingPages = true,
				// "left" means inside and "right" means outside.
				if ((currentAppendix.documentPreferences.facingPages = true) && (i % 2 == 1)) {
					myTextFrame.geometricBounds = [margins.top-8, margins.right, margins.top, docWidth-margins.left];
				} else {
					myTextFrame.geometricBounds = [margins.top-8, margins.left, margins.top, docWidth-margins.right];
				}
				myTextFrame.contents = decodeURI(theFile.name);
				//Create a paragraph style named decodeURI(theFile.name); if 
				//no style by that name already exists.
				try{
					myParagraphStyle = currentAppendix.paragraphStyles.item(paragraphStyleName);
					//If the paragraph style does not exist, trying to get its name will generate an error.
					myName = myParagraphStyle.name;
				}
				catch (myError){
					//The paragraph style did not exist, so create it.
					myParagraphStyle = currentAppendix.paragraphStyles.add({name:paragraphStyleName});
				}
				//At this point, the variable myParagraphStyle contains a reference to a paragraph 
				//style object, which you can now use to specify formatting.
				myTextFrame.parentStory.texts.item(0).applyParagraphStyle(myParagraphStyle, true);
                returnTextPoint = myTextFrame.insertionPoints.firstItem();
				// Start new section
				try {
				var newSection = currentAppendix.sections.add (currentAppendix.pages[i]);
				newSection.marker = decodeURI(theFile.name);
				} catch(e) {
					alert(e);
					alert(i);
					throw(e);
				}
			}
		}
		// Create a temporary text box to place graphic in (to use auto positioning and sizing)
		var TB = currentAppendix.pages[i].textFrames.add({geometricBounds:bounds});
		//decrease the font size of the newly inserted box to 0 to avoid a very misleading "out of pasteboard" error
		//background: if the default font size of the ID document (set by default character style or default paragraph style) causes the text box to overflow it gives you an error saying ("This value would cause one or more objects to leave the pasteboard."). This mainly manifests in pixel based documents as the text box is only 20x20 px large in those cases.
		TB.texts.firstItem().pointSize=1;
		var theRect = TB.insertionPoints.firstItem().rectangles.add();
            theRect.label = "Multi_Page_Importer_Rect";
		// Applying the object style and doing a recompose updates some objects that 
		// the add method doesn't create in the rectangle object
        theRect.appliedObjectStyle = tempObjStyle;
		TB.recompose();
		        
		// Place the current PDF/Ind page into the rectangle object
		try
		{
			var tempGraphic = theRect.place(theFile)[0];
			/* removed 6/25/08
			tempGraphic.graphicLayerOptions.updateLinkOption = (indUpdateType == 0) ?
																							  UpdateLinkOptions.APPLICATION_SETTINGS : 
																							  UpdateLinkOptions.KEEP_OVERRIDES;
			*/
    
			// If all pgs are being added, check that we aren't cruising to the first PDF page again
			if(!noPDFError && !firstTime && tempGraphic.pdfAttributes.pageNumber == 1)
			{
				// If a page was added, nuke it, it's a dupe of the first page
				if(addedAPage)
				{
					currentAppendix.pages[i].remove();
				}
				else
				{
					// Just remove the placed graphic
					TB.remove();
				}
				return;
			}
		}
		catch(e)
		{
            alert("caught: " + e);
			if(e.description.indexOf("Failed to open") != -1 )
			{
				alert("\"" + fileName + "\" doesn't contain a \"" + cropStrings[cropType] + "\" crop type:\n\nPlease try again by selecting a different crop type or open\nthe PDF in Acrobat and perform a \"Save As...\" command.", "PDF Placement Error");
			}
			else
			{
				alert(e);
			}
			if(placeOnLayer)
			{
				currentLayer.remove();
			}
			else
			{
				TB.remove();
			}
			restoreAppendixDefaults(true);
			// Kill the Object style
			tempObjStyle.remove();
			exit(-1);
		}
	
		// Apply any rotation
		if (rotateValues[rotate] != 0) {
			theRect.rotationAngle = rotateValues[rotate];
		}

        //TG: change to fit to page with margins
        theRect.geometricBounds = bounds;
        tempObjStyle.anchoredObjectSettings.anchorXoffset = -margins.left;
        // From documentation:
        // When document.documentPreferences.facingPages = true,
        // "left" means inside and "right" means outside.
        if ((currentAppendix.documentPreferences.facingPages = true) && (i % 2 == 1)) {
            tempObjStyle.anchoredObjectSettings.anchorXoffset = -margins.right;
        }
        tempObjStyle.anchoredObjectSettings.anchorYoffset = margins.top;

        // Fit the placed page according to selected options
        if(keepProp) {
            theRect.fit(FitOptions.proportionally);
            theRect.fit(FitOptions.frameToContent);// Size box down to size of placed page
        } else {
            theRect.fit(FitOptions.contentToFrame);
        }

		// Apply the Object Style to transform the graphic into an anchored item (allows auto positioning)
		theRect.appliedObjectStyle = tempObjStyle;
		
		// Force the text box to reformat itself in order to apply the Object Style
		TB.recompose();
		
		// Release the placed page from the text box and then delete the text box (clean up)
		theRect.anchoredObjectSettings.releaseAnchoredObject();
		TB.remove();
	
		firstTime = false;
        addedAPage = false;
	}
}

// function to restore saved settings back to originals before script ran
// extras parameter is for exiting at different areas of script:
// false: prior to doing anything
// true: end of script or reading PDF file size
function restoreAppendixDefaults(extras)
{
	app.scriptPreferences.userInteractionLevel = oldInteractionPref;
	if(extras == true)
	{
		currentAppendix.zeroPoint = oldZero;
		currentAppendix.viewPreferences.rulerOrigin = oldRulerOrigin;
	}
}


/*********************************************/
/*                                                                */
/*        PDF READER SECTION           */
/*  Extracts count and size of pages    */
/*                                                                */
/********************************************/

// Extract info from the PDF file.
// getSize is a boolean that will also determine page size and rotation of first page
// *** File position changes in this function. ***
// Results are as follows:
// page count = retArray.pgCount
// page width = retArray.pgSize.pgWidth
// page height = retArray.pgSize.pgHeight
function getPDFInfo(theFile, getSize)
{ 
	var flag = 0; // used to keep track if the %EOF line was encountered
	var nlCount = 0; // number of newline characters per line (1 or 2)

	// The array to hold return values
	var retArray = new Array();
	retArray["pgCount"] = -1;
	retArray["pgSize"] = null;

	// Open the PDF file for reading
	theFile.open("r");

    // Search for %EOF line
	// This skips any garbage at the end of the file
	// if FOE% is encountered (%EOF read backwards), flag will be 15
	for(i=0; flag != 15; i++)
	{
		theFile.seek(i,2);
		switch(theFile.readch())
		{
			case "F":
				flag|=1;
				break;
			case "O":
				flag|=2;
				break;
			case "E":
				flag|=4;
				break;
			case "%":
				flag|=8;
				break;
			default:
				flag=0;
				break;
		}
	}
	// Jump back a small distance to allow going forward more easily
	theFile.seek(theFile.tell()-100);

	// Read until startxref section is reached
	while(theFile.readln() != "startxref");

	// Set the position of the first xref section
	var xrefPos = parseInt(theFile.readln(), 10);

	// The array for all the xref sections
	var	xrefArray = new Array();

	// Go to the xref section
	theFile.seek(xrefPos);

	// Determine length of xref entries
	// (not all PDFs are compliant with the requirement of 20 char/entry)
	xrefArray["lineLen"] = determineLineLen(theFile);

	// Get all the xref sections
	while(xrefPos != -1)
	{
		// Go to next section
		theFile.seek(xrefPos);

		// Make sure it's an xref line we went to, otherwise PDF is no good
		if (theFile.readln() != "xref")
		{
			throwError("Cannot determine page count.", true, 99, theFile);
		}

		// Add the current xref section into the main array
		xrefArray[xrefArray.length] = makeXrefEntry(theFile, xrefArray.lineLen);

		// See if there are any more xref sections
		xrefPos = xrefArray[xrefArray.length-1].prevXref;
	}

	// Go get the location of the /Catalog section (the /Root obj)
	var objRef = -1;
	for(i=0; i < xrefArray.length; i++)
	{
		objRef = xrefArray[i].rootObj;
		if(objRef != -1)
		{
			i = xrefArray.length;
		}
	}

	// Double check root obj was found
	if(objRef == -1)
	{
		throwError("Unable to find Root object.", true, 98, theFile);
	}

	// Get the offset of the root section and set file position to it
	var theOffset = getByteOffset(theFile, objRef, xrefArray);
	theFile.seek(theOffset);

	// Determine the obj where the first page is located
	objRef = getRootPageNode(theFile);

	// Get the offset where the root page nod is located and set the file position to it
	theOffset = getByteOffset(theFile, objRef, xrefArray);
	theFile.seek(theOffset);

	// Get the page count info from the root page tree node section
	retArray.pgCount = readPageCount(theFile);

	// Does user need size also? If so, get size info
	if(getSize)
	{
		// Go back to root page tree node
		theFile.seek(theOffset);

		// Flag to tell if page tree root was visited already
		var rootFlag = false;

		// Loop until an actual page obj is found (page tree leaf)
		do
		{
			var getOut = true;

			if(rootFlag)
			{
				// Try to find the line with the /Kids entry
				// Also look for instance when MediBox is in the root obj
				do
				{
					var tempLine = theFile.readln();
				}while(tempLine.indexOf("/Kids") == -1 && tempLine.indexOf(">>") == -1);

			}
			else
			{				
				// Try to first find the line with the /MediaBox entry
				rootFlag = true; // Indicate root page tree was visited
				getOut = false; // Force loop if /MediaBox isn't found here
				do
				{
					var tempLine = theFile.readln();
					if(tempLine.indexOf("/MediaBox") != -1)
					{
						getOut = true;
						break;
					}
				}while(tempLine.indexOf(">>") == -1);

				if(!getOut)
				{
					// Reset the file pointer to the beginning of the root obj again
					theFile.seek(theOffset)
				}
			}

			// If /Kids entry was found, still at an internal page tree node
			if(tempLine.indexOf("/Kids") != -1)
			{
				// Check if the array is on the same line
				if(tempLine.indexOf("R") != -1)
				{
					// Grab the obj ref for the first page
					objRef = parseInt(tempLine.split("/Kids")[1].split("[")[1]);
				}
				else
				{
					// Go down one line
					tempLine = theFile.readln();

					// Check if the opening bracket is on this line
					if(tempLine.indexOf("[") != -1)
					{
						// Grab the obj ref for the first page
						objRef = parseInt(tempLine.split("[")[1]);
					}
					else
					{
						// Grab the obj ref for the first page
						objRef = parseInt(tempLine);
					}

				}

				// Get the file offset for the page obj and set file pos to it
				theOffset = getByteOffset(theFile, objRef, xrefArray);
				theFile.seek(theOffset);
				getOut = false;
			}
		}while(!getOut);

		// Make sure file position is correct if finally at a leaf
		theFile.seek(theOffset);

		// Go get the page sizes
		retArray.pgSize = getPageSize(theFile);
	}

	// Close the PDF file, finally all done!
	theFile.close();

	return retArray;
}

// Function to create an array of xref info
// File position must be set to second line of xref section
// *** File position changes in this function. ***
function makeXrefEntry(theFile, lineLen)
{
	var newEntry = new Array();
	newEntry["theSects"] = new Array();
	var tempLine = theFile.readln();

	// Save info
	newEntry.theSects[0] = makeXrefSection(tempLine, theFile.tell());

	// Try to get to trailer line
	var xrefSec = newEntry.theSects[newEntry.theSects.length-1].refPos;
	var numObjs = newEntry.theSects[newEntry.theSects.length-1].numObjs;
	do
	{
		var getOut = true;
		for(i=0; i<numObjs;i++)
		{
			theFile.readln(); // get past the objects: tell( ) method is all screwed up in CS4
		}
		tempLine = theFile.readln();
		if(tempLine.indexOf("trailer") == -1)
		{ 
			// Found another xref section, create an entry for it
			var tempArray = makeXrefSection(tempLine, theFile.tell());
			newEntry.theSects[newEntry.theSects.length] = tempArray;
			xrefSec = tempArray.refPos;
			numObjs = tempArray.numObjs;
			getOut = false;
		}
	}while(!getOut);

	// Read line with trailer dict info in it
	// Need to get /Root object ref
	newEntry["rootObj"] = -1;
	newEntry["prevXref"] = -1;
	do
	{
		tempLine = theFile.readln();
		if(tempLine.indexOf("/Root") != -1)
		{			
			// Extract the obj location where the root of the page tree is located:
			newEntry.rootObj = parseInt(tempLine.substring(tempLine.indexOf("/Root") + 5), 10);
		}
		if(tempLine.indexOf("/Prev") != -1)
		{
			newEntry.prevXref = parseInt(tempLine.substring(tempLine.indexOf("/Prev") + 5), 10);
		}

	}while(tempLine.indexOf(">>") == -1);

	return newEntry;
}

// Function to save xref info to a given array
function makeXrefSection(theLine, thePos)
{
	var tempArray = new Array();
	var temp = theLine.split(" ");
	tempArray["startObj"] = parseInt(temp[0], 10);
	tempArray["numObjs"] = parseInt(temp[1], 10);
	tempArray["refPos"] = thePos;
	return tempArray;
}

// Function that gets the page count form a root page section
// *** File position changes in this function. ***
function readPageCount(theFile)
{
	// Read in first line of section
	var theLine = theFile.readln();

	// Locate the line containing the /Count entry
	while(theLine.indexOf("/Count") == -1)
	{
		theLine = theFile.readln();
	}

	// Extract the page count
	return parseInt(theLine.substring(theLine.indexOf("/Count") +6), 10);
}

// Function to determine length of xref entries
// Not all PDFs conform to the 20 char/entry requirement
// *** File position changes in this function. ***
function determineLineLen(theFile)
{
	// Skip xref line
	theFile.readln();
	var lineLen = -1;

	// Loop trying to find lineLen
	do
	{
		var getOut = true;
		 var tempLine = theFile.readln();
		if(tempLine != "trailer")
		{
			// Get the number of object enteries in this section
			var numObj = parseInt(tempLine.split(" ")[1]);

			// If there is more than one entry in this section, use them to determime lineLen
			if(numObj > 1)
			{
				theFile.readln();
				var tempPos = theFile.tell();
				theFile.readln();
				lineLen = theFile.tell() - tempPos;
			}
			else
			{
				if(numObj == 1)
				{
					// Skip the single entry
					theFile.readln();
				}
				getOut = false;
			}
		}
		else
		{
			// Read next line(s) and extract previous xref section
			var getOut = false;
			do
			{
				tempLine = theFile.readln();
				if(tempLine.indexOf("/Prev") != -1)
				{
					theFile.seek(parseInt(tempLine.substring(tempLine.indexOf("/Prev") + 5)));
					getOut = true;
				}
			}while(tempLine.indexOf(">>") == -1 && !getOut);
			theFile.readln(); // Skip the xref line
			getOut = false;
		}
	}while(!getOut);

	// Check if there was a problem determining the line length
	if(lineLen == -1)
	{
		throwError("Unable to determine xref dictionary line length.", true, 97, theFile);
	}

	return lineLen;
}

// Function that determines the byte offset of an object number
// Searches the built array of xref sections and reads the offset for theObj
// *** File position changes in this function. ***
function getByteOffset(theFile, theObj, xrefArray)
{
	var theOffset = -1;

	// Look for the theObj in all sections found previously
	for(i = 0; i < xrefArray.length; i++)
	{
		var tempArray = xrefArray[i];
		for(j=0; j < tempArray.theSects.length; j++)
		{
			 var tempArray2 = tempArray.theSects[j];

			// See if theObj falls within this section
			if(tempArray2.startObj <= theObj && theObj <= tempArray2.startObj + tempArray2.numObjs -1)
			{
				theFile.seek((tempArray2.refPos + ((theObj - tempArray2.startObj) * xrefArray.lineLen)));

				// Get the location of the obj
				var tempLine = theFile.readln();

				// Check if this is an old obj, if so ignore it
				// An xref entry with n is live, with f is not
				if(tempLine.indexOf("n") != -1)
				{
					theOffset = parseInt(tempLine, 10);

					// Cleanly get out of both loops
					j = tempArray.theSects.length;
					i = xrefArray.length;
				}
			}
		}
	}

	return theOffset;
}

// Function to extract the root page node object from a section
// File position must be at the start of the root page node
// *** File position changes in this function. ***
function getRootPageNode(theFile)
{
	var tempLine = theFile.readln();

	// Go to line with /Page token in it
	while(tempLine.indexOf("/Pages") == -1)
	{
		tempLine = theFile.readln();
	}

	// Extract the root page obj number
	return parseInt(tempLine.substring(tempLine.indexOf("/Pages") + 6), 10);
}

// Function to extract the sizes from a page reference section
// File position must be at the start of the page object
// *** File position changes in this function. ***
function getPageSize(theFile)
{
	var hasTrimBox = false; // Prevent MediaBox from overwriting TrimBox info
	var charOffset = -1;
	var isRotated = false; // Page rotated 90 or 270 degrees?
	var foundSize = false; // Was a size found?
	do
	{
		var theLine = theFile.readln();
		if(!hasTrimBox && (charOffset = theLine.indexOf("/MediaBox")) != -1)
		{
			// Is the array on the same line?
			if(theLine.indexOf("[", charOffset + 9) == -1)
			{
				// Need to go down one line to find the array
				theLine = theFile.readln();
				// Extract the values of the MediaBox array (x1, y1, x2, y2)
				var theNums = theLine.split("[")[1].split("]")[0].split(" ");
			}
			else
			{
				// Extract the values of the MediaBox array (x1, y1, x2, y2)
				var theNums = theLine.split("/MediaBox")[1].split("[")[1].split("]")[0].split(" ");
			}

			// Take care of leading space
			if(theNums[0] == "")
			{
				theNums = theNums.slice(1);
			}

			foundSize = true;
		}
		if((charOffset = theLine.indexOf("/TrimBox")) != -1)
		{
			// Is the array on the same line?
			if(theLine.indexOf("[", charOffset + 8) == -1)
			{
				// Need to go down one line to find the array
				theLine = theFile.readln();
				// Extract the values of the MediaBox array (x1, y1, x2, y2)
				var theNums = theLine.split("[")[1].split("]")[0].split(" ");
			}
			else
			{
				// Extract the values of the MediaBox array (x1, y1, x2, y2)
				var theNums = theLine.split("/TrimBox")[1].split("[")[1].split("]")[0].split(" ");
			}

			// Prevent MediaBox overwriting TrimBox values
			hasTrimBox = true;

			// Take care of leading space
			if(theNums[0] == "")
			{
				theNums = theNums.slice(1);
			}

			foundSize = true;
		}
		if((charOffset = theLine.indexOf("/Rotate") ) != 1)
		{
			var rotVal = parseInt(theLine.substring(charOffset + 7));
			if(rotVal == 90 || rotVal == 270)
			{
				isRotated = true;
			}
		}
	}while(theLine.indexOf(">>") == -1);

	// Check if a size array wasn't found
	if(!foundSize)
	{
		throwError("Unable to determine PDF page size.", true, 96, theFile);
	}

	// Do the math
	var xSize =	parseFloat(theNums[2]) - parseFloat(theNums[0]);
	var ySize =	parseFloat(theNums[3]) - parseFloat(theNums[1]);

	// One last check that sizes are actually numbers
	if(isNaN(xSize) || isNaN(ySize))
	{
		throwError("One or both page dimensions could not be calculated.", true, 95, theFile);
	}

	// Use rotation to determine orientation of pages
	var ret = new Array();
	ret["width"] = isRotated ? ySize : xSize;
	ret["height"] = isRotated ? xSize : ySize;

	return ret;
}

// Extract info from the document being placed
// Need to open without showing window and then close it
// right after collecting the info
function getINDinfo(theDoc)
{
	// Open it
	var temp = app.open(theDoc, false);
	var placementINFO = new Array();
	var pgSize = new Array();
	// Get info as needed
	placementINFO["pgCount"] = temp.pages.length;
	pgSize["height"] = temp.documentPreferences.pageHeight;
	pgSize["width"] = temp.documentPreferences.pageWidth;
	placementINFO["vUnits"] = temp.viewPreferences.verticalMeasurementUnits;
	placementINFO["hUnits"] = temp.viewPreferences.horizontalMeasurementUnits;
	placementINFO["pgSize"] = pgSize;
	// Close the document
	temp.close(SaveOptions.NO);
	return placementINFO;
}

// File filter for the mac to only show indy and pdf files
function macFileFilter(fileToTest)
{ 
	if(fileToTest.name.indexOf(".pdf") > -1 || fileToTest.name.indexOf(".ai") > -1 || fileToTest.name.indexOf(".ind") > -1)
		return true; 		 
	else
		return false;	 
}

