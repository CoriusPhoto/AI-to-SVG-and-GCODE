//*******************************************************
// Diecut_To_SVGandGCode.jsx
// Version 1.0
//
// Copyright 2023 Corius
// Comments or suggestions to contact@corius.fr
//
//*******************************************************

// Script version
var Diecut_To_SVGandGCode = 'v1.0';

// CONSOLE BLANK LINE
//$.writeln('          ----------<<<<<<<<<<     EXEC START     >>>>>>>>>>----------');

// AI document variables
var docObj = app.activeDocument;
var docName = docObj.fullName;
var docFolder = docName.parent.fsName;

// Data array variables
var visibleLayersData = []; // [LayerName, LayerID, PathData[]]
var visibleLayersDistanceData = []; // [LayerName, LayerID, PathDistanceData[]]

var allPathsByLayer = new Array();  // allPathsByLayer[i][LayerName, allPaths[]]
var allOrderedPathsByLayer = new Array();  // allOrderedPathsByLayer[i][LayerName, allOrderedPaths[]]
var allConvertedPathByLayer = new Array();  // allConvertedPathByLayer[i][LayerName, allSVG[], allGCode[]]

var origin = [docObj.artboards[0].artboardRect[0],docObj.artboards[0].artboardRect[3]];
var ORIGIN = new Point();
var ENDPOINT = new Point();
ORIGIN.x = docObj.artboards[0].artboardRect[0];
ORIGIN.y = docObj.artboards[0].artboardRect[3];

var unitTable = new Array();
unitTable[0] = new Array('Full text', 'Unknown', 'Inches', 'Centimeters', 'Points', 'Picas', 'Millimeters', 'Qs', 'Pixels', 'FeetInches', 'Meters', 'Yards', 'Feet');
unitTable[1] = new Array('mm equivalent', 'Unknown', 25.4, 10, null, null, 1, null, 25.4/72, null, 1000, 914.4, 304.8);

// SVG variables
var unitRatio = null;
var svgRatio = 25.4/72;
var GLOBALSVG = true; // set to true to use global coordinates for SVG conversion, set to false to use local coodinates (local coordinates not yet implemented). 
var C = GLOBALSVG ? ' C' : ' c';
var S = GLOBALSVG ? ' S' : ' s';
var A = GLOBALSVG ? ' A' : ' a';
var L = GLOBALSVG ? ' L' : ' l';
var V = GLOBALSVG ? ' V' : ' v';
var H = GLOBALSVG ? ' H' : ' h';
var S = GLOBALSVG ? ' S' : ' s';
var PRECISION = 1000; // the coordinates precision, rounded to 1/PRECISIONth (so 1000 means up to 3 digits after the dot, so 1/1000th)
var STROKE_COLORS = new Array();
var SVGDEFSTRING = '';

// LASER GRBL VARIABLES
var PARAMS = '';
var CUT_POWER = 0;
var CUT_SPEED = 0;
var CUT_CYCLES = 0;
var FOLD_POWER = 0;
var FOLD_SPEED = 0;
var FOLD_CYCLES = 0;
var TEST_POWER = 0;
var TEST_SPEED = 0;
var TEST_CYCLES = 0;
var FULL_GCODE = 'G90\nM3 S0\n';
var CUT_GCODE = '';
var FOLD_GCODE = '';
var CUT_SVG = '';
var FOLD_SVG = '';
var NEED_FOLD = false;

var POWER;
var SPEED;
var CYCLES;
var lastX = 0;
var lastY = 0;
var firstX = 0;
var firstY = 0;
var TRIMORIGINY = 0;

var BEZIERSTEPS = 100;

// File variables
var svgFolder = '\\SVG';
var gcodeFolder = '\\GCODE';
var MergeSVG = false;
var MergeGCODE = true;
var ExportedFiles = '';

getData();

//$.writeln('visibleLayersData : '+visibleLayersData); 
//$.writeln('origin : '+origin); 
//calculDistance(origin,[-1,100,-350]);

function getData(){
    var myLayer;
    
    var svgxmlstring = '<?xml version="1.0" encoding="UTF-8"?>\n<svg ';
    var svgendheadertxt = '" xmlns="http://www.w3.org/2000/svg"';
    var id;
    var w = Math.round(100 * docObj.width)/100;
    var h = Math.round(100 * docObj.height)/100;
    var j;
    var myPath;
    var myGroup;
    
    var docUnit = new String(docObj.rulerUnits);
    docUnit = docUnit.split('.')[1];
    for (var i=0;i<unitTable[0].length && unitRatio == null;i++){
        if (unitTable[0][i] == docUnit){ 
            if (i > 1 && unitTable[1][i] != null){
                unitRatio = unitTable[1][i];
                svgRatio = unitRatio * svgRatio;
            } else {
                $.writeln('ERROR - can\'t process document unit : '+docObj.rulerUnits); 
            }
        }  /*else {
            //$.writeln('next i'); 
        }*/
    }
    //$.writeln('unitRatio : '+unitRatio);
    //$.writeln('svgRatio : '+svgRatio);
    
    //svgendheadertxt += ' width="' +Math.round(PRECISION * w * svgRatio)/PRECISION+ 'mm" height="' +Math.round(PRECISION * h * svgRatio)/PRECISION+ 'mm" viewBox="0 0 ' +w+ ' ' +h+'">'
    
    TRIMORIGINY = Math.round(PRECISION * h * svgRatio)/PRECISION;
    
    for(var i=0; i<docObj.layers.length; i++){
        myLayer = docObj.layers[i];
        
        if (myLayer.visible && myLayer.name != 'GCODE_PARAMS' && myLayer.name != 'Margein' && myLayer.name != 'draft'){
            j = allPathsByLayer.length;
            allPathsByLayer[j] = new Array();
            allPathsByLayer[j][0] = myLayer.name;
            allPathsByLayer[j][1] = new Array();
            allOrderedPathsByLayer[j] = new Array();
            allOrderedPathsByLayer[j][0] = myLayer.name;
            allOrderedPathsByLayer[j][1] = new Array();
            allOrderedPathsByLayer[j][2] = new Array(); //the [2] will store all the "direction" to use current or inversed path direction
            allConvertedPathByLayer[j] = new Array();
            allConvertedPathByLayer[j][0] = myLayer.name;
            allConvertedPathByLayer[j][1] = new Array(); // to store SVG path full XML sentence "<path class=… d=…/>"
            allConvertedPathByLayer[j][2] = new Array(); // to store complete GCode for each path
            STROKE_COLORS[j] = new Array();
            STROKE_COLORS[j][0] = myLayer.name;
            
            for(var k=0; k < myLayer.pathItems.length; k++){
                myPath = myLayer.pathItems[k];
                if (myPath.hidden == false){
                    allPathsByLayer[j][1].push(myPath);
                }
            }
            for(k=0; k < myLayer.groupItems.length; k++){
                myGroup = myLayer.groupItems[k]; 
                if (myGroup.hidden == false){
                    pathInGroup(myGroup,j);
                }
            }
            if (myLayer.name == 'FOLD' && allPathsByLayer[j][1].length > 0){
                NEED_FOLD = true;
                //$.writeln('NEED_FOLD : '+NEED_FOLD);
            }
        } /*else {
            //$.writeln('invisible layer : '+myLayer.name+'  ID : '+i); 
        }*/
    }

    getSettings();
    
    getAllLayersPathsOrder();
    
    getAllSVGPaths();
    
    getAllGCODEPaths();
    
    alert('Done ! \nThe following files have been exported : '+ExportedFiles, 'Script complete');
}

function pathInGroup(myGroup,myID){
    var myPath;
    var myItem;
    
    for (var i=0; i<myGroup.pageItems.length;i++){
        myItem = myGroup.pageItems[i];
        if (myItem.typename == 'PathItem' && myItem.hidden == false){
            myPath = myItem;
            allPathsByLayer[myID][1].push(myPath);
        } else if (myItem.typename == 'GroupItem' && myItem.hidden == false){
            pathInGroup(myItem,myID);
        }
    }    
}

function getSettings(){
    var myLayer = docObj.layers.getByName('GCODE_PARAMS');
    var myGroup;
    var myPath;
    var layerName;
    var myColor;
    var merge;
    
    
    for (var i=0; i<myLayer.pageItems.length && myGroup==null;i++){
        if (!myLayer.pageItems[i].hidden && myLayer.pageItems[i].typename == 'GroupItem' && myLayer.pageItems[i].name != 'TABLE'){
            myGroup = myLayer.pageItems[i];
        }
    }
    PARAMS = myGroup.name;
    //$.writeln('PARAMS : '+PARAMS); 
    CUT_POWER = 10 * myGroup.textFrames.getByName('CUT_POWER').contents;
    CUT_SPEED = myGroup.textFrames.getByName('CUT_SPEED').contents;
    CUT_CYCLES = myGroup.textFrames.getByName('CUT_CYCLES').contents;
    FOLD_POWER = 10 * myGroup.textFrames.getByName('FOLD_POWER').contents;
    FOLD_SPEED = myGroup.textFrames.getByName('FOLD_SPEED').contents;
    FOLD_CYCLES = myGroup.textFrames.getByName('FOLD_CYCLES').contents;
    TEST_POWER = 10 * myGroup.textFrames.getByName('TEST_POWER').contents;
    TEST_SPEED = myGroup.textFrames.getByName('TEST_SPEED').contents;
    TEST_CYCLES = myGroup.textFrames.getByName('TEST_CYCLES').contents;   
    
    for (i=0;i<STROKE_COLORS.length;i++){
        layerName = STROKE_COLORS[i][0];
        myPath = myLayer.groupItems.getByName('TABLE').pathItems.getByName(layerName+'_STROKE');
        myColor = myPath.strokeColor;
        STROKE_COLORS[i][1] = cmykToHex(myColor.cyan, myColor.magenta, myColor.yellow, myColor.black);
        //$.writeln('LAYER '+layerName+' stroke color = ['+myColor.cyan+']['+myColor.magenta+']['+myColor.yellow+']['+myColor.black+']');
        //$.writeln('LAYER '+layerName+' stroke color = '+STROKE_COLORS[i][1]);
        SVGDEFSTRING += '<defs>\n<style>\n.'+layerName+' {\nfill: none;\nstroke: '+STROKE_COLORS[i][1]+';\nstroke-miterlimit: 10;\nstroke-width: .5px;\n}\n</style>\n</defs>\n';
    }

    BEZIERSTEPS = myLayer.groupItems.getByName('TABLE').textFrames.getByName('BEZIER').contents;
    PRECISION = myLayer.groupItems.getByName('TABLE').textFrames.getByName('PRECISION').contents;
    
    merge = myLayer.groupItems.getByName('TABLE').textFrames.getByName('SVG_MERGE').contents;
    MergeSVG = (merge == 'YES' || merge == 'yes' || merge == 'true' || merge == 'TRUE')? true : false;
    merge = myLayer.groupItems.getByName('TABLE').textFrames.getByName('GCODE_MERGE').contents;
    MergeGCODE = (merge == 'YES' || merge == 'yes' || merge == 'true' || merge == 'TRUE')? true : false;
    
    /*$.writeln('BEZIERSTEPS : '+BEZIERSTEPS);
    $.writeln('PRECISION : '+PRECISION);
    $.writeln('MergeSVG : '+MergeSVG);
    $.writeln('MergeGCODE : '+MergeGCODE);*/
    //myLayer.visible = false;
}

function cmykToHex(c,m,y,k,) {
    var hex,
        rgb;
    //convert cmyk to rgb first
    rgb = cmykToRgb(c,m,y,k,false);
    //then convert rgb to hex
    hex = rgbToHex(rgb.r, rgb.g, rgb.b);
    //return hex color format
    return hex;
}

function cmykToRgb(c, m, y, k, normalized){
    c = (c / 100);
    m = (m / 100);
    y = (y / 100);
    k = (k / 100);
    
    c = c * (1 - k) + k;
    m = m * (1 - k) + k;
    y = y * (1 - k) + k;
    
    var r = 1 - c;
    var g = 1 - m;
    var b = 1 - y;
    
    if(!normalized){
        r = Math.round(255 * r);
        g = Math.round(255 * g);
        b = Math.round(255 * b);
    }
    
    return {
        r: r,
        g: g,
        b: b
    }
}

function rgbToHex(r, g, b) {
    return "#" + componentToHex(r) + componentToHex(g) + componentToHex(b);
}

function componentToHex(c) {
    var hex = c.toString(16);
    return hex.length === 1 ? "0" + hex : hex;
}

function getAllLayersPathsOrder(){
    var myLayerName;
    for (var i=0; i<allPathsByLayer.length; i++){
        getLayerPathsOrder(i); 
        //$.writeln('allOrderedPathsByLayer[i][1].length : '+allOrderedPathsByLayer[i][1].length);
    }
    //$.writeln('all order OK');
}

function getLayerPathsOrder(layerID){
    var startPoint;
    var myArr;
    var endPoint = new Array();
    var closestPathID;
    var closestPath;
    var direction;
    var myPathsArr = allPathsByLayer[layerID][1];
    var cycleNumber = myPathsArr.length;
    var myOrderedPathPoints;
    
    //$.writeln('cycleNumber : '+cycleNumber);
    for (var i=0; i<cycleNumber; i++){
        if (i==0){
            startPoint = new Point();
            startPoint.x = ORIGIN.x;
            startPoint.y = ORIGIN.y;
        } else {
            startPoint = ENDPOINT;
        }
        myArr = findClosestPoint(startPoint,myPathsArr);
        closestPathID = myArr[0];
        direction = myArr[1];
        var lastID = allOrderedPathsByLayer[layerID][1].length;
        var pathToSaveArr = allPathsByLayer[layerID][1].splice(closestPathID,1);
        allOrderedPathsByLayer[layerID][1][lastID] = pathToSaveArr[0];
        allOrderedPathsByLayer[layerID][2][lastID] = direction;
    }
}

function findClosestPoint(startPoint,myPathsArr){
    var myPath;
    var myPoint1;
    var myPoint2;
    var endPoint;
    var dist1;
    var dist2;
    var distance = 100000000;
    var direction = 1;
    var closestPathID = 0;
    
    for(var i=0; i<myPathsArr.length; i++){
        myPath = myPathsArr[i];
        myPoint1 = new Point ();
        myPoint1.x = myPath.pathPoints[0].anchor[0];
        myPoint1.y = myPath.pathPoints[0].anchor[1];
        myPoint2 = new Point ();
        myPoint2.x = myPath.pathPoints[myPath.pathPoints.length -1].anchor[0];
        myPoint2.y = myPath.pathPoints[myPath.pathPoints.length -1].anchor[1];
        dist1 = calculDistance(startPoint,myPoint1);
        dist2 = calculDistance(startPoint,myPoint2);
        
        if (dist1 < distance){
            distance = dist1;
            direction = 1;
            closestPathID = i;
            ENDPOINT.x = myPoint2.x;
            ENDPOINT.y = myPoint2.y;
        }
        if (dist2 < distance){
            distance = dist2;
            direction = -1;
            closestPathID = i;
            ENDPOINT.x = myPoint2.x;
            ENDPOINT.y = myPoint2.y;
        }
    }

    return new Array(closestPathID,direction);
}

function calculDistance(point1,point2){
    var dist = 0;
    var Xdif = point1.x - point2.x;
    var Ydif = point1.y - point2.y;
    
    dist = Math.abs(Xdif) + Math.abs(Ydif);
    
    return dist;
}

function getAllSVGPaths(){
    var myPathsSVGtxt;
    var myPathsSVGarr;
    for (var i=0; i<allOrderedPathsByLayer.length; i++){
        if (allOrderedPathsByLayer[i][1] != null && allOrderedPathsByLayer[i][1].length >0){
            myPathsSVGarr = getSVGPaths(i);
            allConvertedPathByLayer[i][1] = myPathsSVGarr;
            makeSVGfile(i);
        }
     }
}

function getAllGCODEPaths(){
    var myPathsGCODEtxt;
    var myPathsGCODEarr;
    for (var i=0; i<allOrderedPathsByLayer.length; i++){
        if (allOrderedPathsByLayer[i][1] != null && allOrderedPathsByLayer[i][1].length >0){
            myPathsGCODEarr = getGCODEPaths(i);
            allConvertedPathByLayer[i][2] = myPathsGCODEarr;
            makeGCODEfile(i);
        }
     }
    //$.writeln('all GCODE OK'); 
}

function getSVGPaths(LayerID){
    var myPathsArr = allOrderedPathsByLayer[LayerID][1];
    var myDirectionArr = allOrderedPathsByLayer[LayerID][2];
    var svgTXT = '';
    var svgArr = new Array();
    var layerName =  allOrderedPathsByLayer[LayerID][0];
    
    var myPath;
    var mySegPoint1;
    var mySegPoint2;
    var myDirection;
    var totalLoop;
    var loopCount;
    var previousCurved = false;
    var pathID;
    var startPathPointID;
    var myPointID;
    var curveSection1 = 1; // -1 for half curve before point, 0 for full curve, 1 for half curve after point
    var curveSection2 = -1; // -1 for half curve before point, 0 for full curve, 1 for half curve after point
    var P1id;
    var P2id;
    var anteP1 = new Array();
    var P1 = new Array();
    var postP1 = new Array();
    var anteP2 = new Array();
    var P2 = new Array();
    var postP2 = new Array();
    
    //$.writeln('myPathsArr.length : '+myPathsArr.length);
    for (pathID=0; pathID<myPathsArr.length;pathID++){
        svgTXT = '<path class="'+layerName+'" d="M';
        myPath = myPathsArr[pathID];
        myDirection = myDirectionArr[pathID];
        previousCurved = false;
        
        totalLoop = myPath.closed ? myPath.pathPoints.length + 1 : myPath.pathPoints.length;
        startPathPointID = (myDirection > 0) ? 0 : myPath.pathPoints.length - 1;
        for (loopCount = 0 ; loopCount < myPath.pathPoints.length ; loopCount++){
            myPointID = startPathPointID + loopCount * myDirection;
            P1id = myPointID;
            mySegPoint1 = myPath.pathPoints[myPointID];
            if (loopCount < myPath.pathPoints.length - 1) {
                mySegPoint2 = myPath.pathPoints[myPointID + myDirection];
                P2id = myPointID + myDirection;
            } else {
                mySegPoint2 = (myDirection > 0) ? myPath.pathPoints[0] : myPath.pathPoints[myPath.pathPoints.length - 1];
                P2id = (myDirection > 0) ? 0 : myPath.pathPoints.length - 1;
            }
        
            anteP1[0] = (myDirection > 0) ? roundDecimal(mySegPoint1.leftDirection[0]) : roundDecimal(mySegPoint1.rightDirection[0]);
            anteP1[1] = (myDirection > 0) ? -roundDecimal(mySegPoint1.leftDirection[1]) : -roundDecimal(mySegPoint1.rightDirection[1]);
            P1[0] = roundDecimal(mySegPoint1.anchor[0]);
            P1[1] = -roundDecimal(mySegPoint1.anchor[1]);
            postP1[0] = (myDirection > 0) ? roundDecimal(mySegPoint1.rightDirection[0]) : roundDecimal(mySegPoint1.leftDirection[0]);
            postP1[1] = (myDirection > 0) ? -roundDecimal(mySegPoint1.rightDirection[1]) : -roundDecimal(mySegPoint1.leftDirection[1]);
            
            anteP2[0] = (myDirection > 0) ? roundDecimal(mySegPoint2.leftDirection[0]) : roundDecimal(mySegPoint2.rightDirection[0]);
            anteP2[1] = (myDirection > 0) ? -roundDecimal(mySegPoint2.leftDirection[1]) : -roundDecimal(mySegPoint2.rightDirection[1]);
            P2[0] = roundDecimal(mySegPoint2.anchor[0]);
            P2[1] = -roundDecimal(mySegPoint2.anchor[1]);
            postP2[0] = (myDirection > 0) ? roundDecimal(mySegPoint2.rightDirection[0]) : roundDecimal(mySegPoint2.leftDirection[0]);
            postP2[1] = (myDirection > 0) ? -roundDecimal(mySegPoint2.rightDirection[1]) : -roundDecimal(mySegPoint2.leftDirection[1]);;
            
            if (loopCount == 0) {
                svgTXT += P1[0] + ',' + P1[1];
            }
            if (loopCount < myPath.pathPoints.length - 1){
                if (isCurved(mySegPoint1, curveSection1 * myDirection)){
                    svgTXT += !previousCurved ? C : ',';
                    svgTXT += postP1[0] + ',' + postP1[1] + ','  + anteP2[0] + ',' + anteP2[1] + ','  + P2[0] + ',' + P2[1];
                    previousCurved = true;
                } else if (isCurved(mySegPoint2, curveSection2 * myDirection)){
                    svgTXT += S;
                    svgTXT += anteP2[0] + ',' + anteP2[1] + ','  + P2[0] + ',' + P2[1];
                    previousCurved = false;
                } else if (P1[0] == P2[0]){
                    svgTXT += V;
                    svgTXT += P2[1];
                    previousCurved = false;
                } else if (P1[1] == P2[1]){
                    svgTXT += H;
                    svgTXT += P2[0];
                    previousCurved = false;
                } else {
                    svgTXT += L;
                    svgTXT += P2[0] + ',' + P2[1];
                    previousCurved = false                }
            } else if(myPath.closed){
                if (isCurved(mySegPoint2, curveSection2 * myDirection)){
                    svgTXT += !previousCurved ? C : ',';
                    svgTXT += postP1[0] + ',' + postP1[1] + ','  + anteP2[0] + ',' + anteP2[1] + ','  + P2[0] + ',' + P2[1];
                    previousCurved = true;
                } else if (isCurved(mySegPoint1, curveSection1 * myDirection)){
                    // THIS ELSE-IF SEEMS TO BE NEVER USED
                    //$.writeln('Adding final S part');
                    svgTXT += S;
                    svgTXT += anteP1[0] + ',' + anteP1[1];
                    previousCurved = false;
                } else {
                    svgTXT += ' Z';
                }
            }
        }
        svgTXT += '"/>\n';
        svgArr.push(svgTXT);
    }
    
    return svgArr;
}

function isCurved(myPathPoint, section){
    var myX = roundDecimal(myPathPoint.anchor[0]);
    var myY = -roundDecimal(myPathPoint.anchor[1]);
    var myLeftX = roundDecimal(myPathPoint.leftDirection[0]);
    var myLeftY = -roundDecimal(myPathPoint.leftDirection[1]);
    var myRightX = roundDecimal(myPathPoint.rightDirection[0]);
    var myRightY = -roundDecimal(myPathPoint.rightDirection[1]);
    
    var curved = false;
    if (section == 0 && (myLeftX != myX || myRightX != myX || myLeftY != myY || myRightY != myY)){
        curved = true;
    } else if (section > 0){
        curved = (myRightX != myX || myRightY != myY) ? true : false;
    } else if (section < 0){
        curved = (myLeftX != myX || myLeftY != myY) ? true : false;
    }
    
    return curved;
}

function roundDecimal(num){
    num = Math.round(PRECISION * num);
    num /= PRECISION;
    
    return num;
}

function makeSVGfile(layerID){
    var layerName =  allOrderedPathsByLayer[layerID][0];
    var myFileName;
    var myFile;
    var myFolderName = docFolder + svgFolder;
    var myFolder = new Folder(myFolderName);
    var strokeColor = STROKE_COLORS[layerID][1];
    var svgxmlstring = '<?xml version="1.0" encoding="UTF-8"?>\n<svg  id="CoriusSVG-GCODE" data-name="'+layerName+'" ';
    var svgendheadertxt = 'xmlns="http://www.w3.org/2000/svg"';
    var id;
    var w = Math.round(100 * docObj.width)/100;
    var h = Math.round(100 * docObj.height)/100;
    var j;
    var myPathsString = allConvertedPathByLayer[layerID][1].join('');
    var svgDefstring = '<defs>\n<style>\n.'+layerName+' {\nfill: none;\nstroke: '+strokeColor+';\nstroke-miterlimit: 10;\nstroke-width: .5px;\n}\n</style>\n</defs>\n';
        
    svgendheadertxt += ' width="' +Math.round(PRECISION * w * svgRatio)/PRECISION+ 'mm" height="' +Math.round(PRECISION * h * svgRatio)/PRECISION+ 'mm" viewBox="0 0 ' +w+ ' ' +h+'">\n'; 
    
    if (layerName == 'TEST' || !NEED_FOLD || !MergeSVG){
        myFileName = docObj.name.substring(0,docObj.name.lastIndexOf ('.ai'))+'_'+PARAMS+'_'+layerName+'.svg';   
        svgxmlstring += svgendheadertxt + svgDefstring + myPathsString + '</svg>';
        
        myFolder.create();
        
        myFile = new File(myFolderName+'\\'+myFileName);
        myFile.encoding = "BINARY";
        myFile.open('w');
        myFile.write(svgxmlstring);
        myFile.close();
        ExportedFiles += '\n'+myFileName;
    } else {
        if (layerName == 'CUT'){
            CUT_SVG = myPathsString;
        }
        if (layerName == 'FOLD'){
            FOLD_SVG = myPathsString;
        }
        if (CUT_SVG != '' && FOLD_SVG != ''){        
            svgxmlstring += svgendheadertxt + SVGDEFSTRING + CUT_SVG + FOLD_SVG + '</svg>';
            myFileName = docObj.name.substring(0,docObj.name.lastIndexOf ('.ai'))+'_'+PARAMS+'_CUT+FOLD'+'.svg';
            myFolder.create();
            
            myFile = new File(myFolderName+'\\'+myFileName);
            myFile.encoding = "BINARY";
            myFile.open('w');
            myFile.write(svgxmlstring);
            myFile.close();
            ExportedFiles += '\n'+myFileName;
        }
    }
}

function makeGCODEfile(layerID){
    var layerName =  allOrderedPathsByLayer[layerID][0];
    var myFileName;
    var myFile;
    var myFolderName = docFolder + gcodeFolder;
    var myFolder = new Folder(myFolderName);
    var gcodestring = '';
    var gcodeEnd = 'S0\nM5 S0\nG0 X0 Y0 Z0';
    
    for (var i=0; i<CYCLES; i++){
        gcodestring += allConvertedPathByLayer[layerID][2].join('');
    }
    
    if (layerName == 'TEST' || !NEED_FOLD || !MergeGCODE){
        myFileName = docObj.name.substring(0,docObj.name.lastIndexOf ('.ai'))+'_'+PARAMS+'_'+layerName+'.nc';
        myFolder.create();
        
        myFile = new File(myFolderName+'\\'+myFileName);
        myFile.encoding = "BINARY";
        myFile.open('w');
        myFile.write(FULL_GCODE+gcodestring+gcodeEnd);
        myFile.close();
        ExportedFiles += '\n'+myFileName;
    } else {
        if (layerName == 'CUT'){
            CUT_GCODE = gcodestring;
        }
        if (layerName == 'FOLD'){
            FOLD_GCODE = gcodestring;
        }
        if (CUT_GCODE != '' && FOLD_GCODE != ''){            
            myFileName = docObj.name.substring(0,docObj.name.lastIndexOf ('.ai'))+'_'+PARAMS+'_CUT+FOLD'+'.nc';
            myFolder.create();
            
            myFile = new File(myFolderName+'\\'+myFileName);
            myFile.encoding = "BINARY";
            myFile.open('w');
            myFile.write(FULL_GCODE+FOLD_GCODE+CUT_GCODE+gcodeEnd);
            myFile.close();
            ExportedFiles += '\n'+myFileName;
        }
    }
}


function getGCODEPaths(LayerID){
    var mySvgPathsArr = allConvertedPathByLayer[LayerID][1];
    var myGcodeArr = allConvertedPathByLayer[LayerID][2];
    var layerName =  allConvertedPathByLayer[LayerID][0];
    var myPathXML;
    var myPath;
    var myPathCommandsArr = new Array();
    var myGcode = 'S0\n';
    var myCommandArr;
    var mySplittedCommandArr;
    
    if(layerName == 'CUT'){
        POWER = CUT_POWER;
        SPEED = CUT_SPEED;
        CYCLES = CUT_CYCLES;
    } else if(layerName == 'FOLD'){
        POWER = FOLD_POWER;
        SPEED = FOLD_SPEED;
        CYCLES = FOLD_CYCLES;
    } else if(layerName == 'TEST'){
        POWER = TEST_POWER;
        SPEED = TEST_SPEED;
        CYCLES = TEST_CYCLES;
    }
    for (var i=0 ; i<mySvgPathsArr.length;i++){
        myPathXML = new XML(mySvgPathsArr[i]);
        myPath = myPathXML.attribute ('d').toString();
        mySplittedCommandArr = splitCommands(myPath);
        for (var j=0; j<mySplittedCommandArr.length; j++){
            myPathCommandsArr.push(mySplittedCommandArr[j]);
        }
    }
    
    myGcodeArr.push(myGcode);
    for (i=0 ; i<myPathCommandsArr.length;i++){
        myCommandArr = myPathCommandsArr[i];
        myGcode = convertPathToGCode(myCommandArr,SPEED,POWER);
        myGcodeArr.push(myGcode);
    }
    
    return myGcodeArr;
}



function splitCommands(pathData) {
    var lastX = 0;
    var lastY = 0;
    
    var commands = new Array();
    var startIndex = 0;
    var endIndex;
    var commandIndex = -1;
    var relicat;
    var relicatSplitedXY;
    
    for (endIndex = 0; endIndex < pathData.length; endIndex++){
        if(pathData.charAt(endIndex).match(new RegExp('[MLACVHZS]','i')) != null || endIndex == pathData.length - 1){
            if (startIndex > 0){
                if (pathData.charAt(endIndex-1) == ' '){
                    relicat = pathData.substring(startIndex,endIndex-1).split(',');
                } else if(endIndex == pathData.length - 1 && pathData.charAt(endIndex).match(new RegExp('[MLACVHZS]','i')) == null){
                    relicat = pathData.substring(startIndex).split(',');
                } else {
                    relicat = pathData.substring(startIndex,endIndex).split(',');
                }
                for (var i=0; i<relicat.length; i++){
                    commands[commandIndex].push(relicat[i]);
                }
            }
            if (endIndex != pathData.length - 1 || (endIndex == pathData.length - 1 && pathData.charAt(endIndex).match(new RegExp('[Z]','i')) != null)){
                commandIndex += 1;
                commands.push(new Array());
                commands[commandIndex].push(new Array());
                commands[commandIndex][0] = pathData.charAt(endIndex);            
                startIndex = endIndex + 1;
            }
        }
    }
    return commands;
}

function convertPathToGCode(commands,speed,power) {
    var gcode = '';
    var codeLine = '';
    var trimX = 0;
    var trimY = 0;
    var myCommand;
    var j;
    var myX;
    var myY;
    var laserON = false;
    var feedStr;
    var tan1X;
    var tan1Y;
    var tan2X;
    var tan2Y;
    var p2X;
    var p2Y;
    var curveCoordinates;

    // Conversion des commandes du chemin en G-code
    myCommand = commands[0];
    //$.writeln('The full command to convert is : '+commands);
    if (myCommand == 'M' || myCommand == 'm'){
        for (j=1; j<commands.length; j++){
            if (j==1){
                laserON = false;
                gcode += 'S0 \nG0';
                feedStr = '';
            } else {
                laserON = true;
                gcode += 'S'+power+' \nG1';
                feedStr = 'F'+speed;
            }
            myX = parseFloat(commands[j]);
            j +=1;
            myY = parseFloat(commands[j]);
            if (j==2){
                firstX = myX;
                firstY = myY;
            }
            if (myCommand == 'm'){
                trimX = lastX;
                trimY = lastY;
            } else {
                trimX = 0;
                trimY = 0;
            }
            codeLine = getMoveToGCode(myX,myY,trimX,trimY,feedStr);
            gcode += codeLine;
            lastX = myX;
            lastY = myY;
        }
    } else if (myCommand == 'H' || myCommand == 'h'){
        // Conversion du déplacement linéaire en G-code
        myX = parseFloat(commands[1]);
        gcode += 'S'+power+' \nG1';
        feedStr = 'F'+speed;
        if (myCommand == 'h'){
            trimX = lastX;
            //trimY = lastY;
        } else {
            trimX = 0;
            //trimY = 0;
        }
        codeLine = getHMoveToGCode(myX,trimX,feedStr);
        gcode += codeLine;
        lastX = myX;
    } else if (myCommand == 'V' || myCommand == 'v'){
        // Conversion du déplacement linéaire en G-code
        myY = parseFloat(commands[1]);
        gcode += 'S'+power+' \nG1';
        feedStr = 'F'+speed;
        if (myCommand == 'v'){
            //trimX = lastX;
            trimY = lastY;
        } else {
            //trimX = 0;
            trimY = 0;
        }
        codeLine = getVMoveToGCode(myY,trimY,feedStr);
        gcode += codeLine;
        lastY= myY;
    } else if (myCommand == 'L' || myCommand == 'l'){
        // Conversion du déplacement linéaire en G-code
        for (j=1; j<commands.length; j++){
            myX = parseFloat(commands[j]);
            j +=1;
            myY = parseFloat(commands[j]);
            gcode += 'S'+power+' \nG1';
            feedStr = 'F'+speed;
            if (myCommand == 'l'){
                trimX = lastX;
                trimY = lastY;
            } else {
                trimX = 0;
                trimY = 0;
            }
            codeLine = getMoveToGCode(myX,myY,trimX,trimY,feedStr);
            gcode += codeLine;
            lastX = myX;
            lastY = myY;
        }
    } else if (myCommand == 'C' || myCommand == 'c'){
        for (j=1; j<commands.length; j++){
            tan1X = parseFloat(commands[j]);
            j +=1;
            tan1Y = parseFloat(commands[j]);
            j +=1;
            tan2X = parseFloat(commands[j]);
            j +=1;
            tan2Y = parseFloat(commands[j]);
            j +=1;
            p2X = parseFloat(commands[j]);
            j +=1;
            p2Y = parseFloat(commands[j]);            
            gcode += 'S'+power+' \nG1';
            feedStr = 'F'+speed;
            curveCoordinates = createCurveCoordinates(BEZIERSTEPS, lastX, lastY, tan1X, tan1Y, tan2X, tan2Y, p2X, p2Y);
            codeLine = '';
            for (var k=0; k<curveCoordinates.length; k++){
                if (myCommand == 'c'){
                    trimX = lastX;
                    trimY = lastY;
                } else {
                    trimX = 0;
                    trimY = 0;
                }
                myX = curveCoordinates[k][0];
                myY = curveCoordinates[k][1];
                codeLine += getMoveToGCode(myX,myY,trimX,trimY,feedStr);
                lastX = myX;
                lastY = myY;
            }
            gcode += codeLine;
            lastX = myX;
            lastY= myY;
        }
    } else if (myCommand == 'Z' || myCommand == 'z'){
        gcode += 'S'+power+' \nG1';
        feedStr = 'F'+speed;
        if (myCommand == 'l'){
            trimX = lastX;
            trimY = lastY;
        } else {
            trimX = 0;
            trimY = 0;
        }
        codeLine = getMoveToGCode(firstX,firstY,trimX,trimY,feedStr);
        gcode += codeLine;
        lastX = myX;
        lastY= myY;
    } else if (myCommand == 'S' || myCommand == 's'){
        for (j=1; j<commands.length; j++){
            tan1X = lastX;
            tan1Y = lastY;
            tan2X = parseFloat(commands[j]);
            j +=1;
            tan2Y = parseFloat(commands[j]);
            j +=1;
            p2X = parseFloat(commands[j]);
            j +=1;
            p2Y = parseFloat(commands[j]);
            gcode += 'S'+power+' \nG1';
            feedStr = 'F'+speed;
            curveCoordinates = createCurveCoordinates(BEZIERSTEPS, lastX, lastY, tan1X, tan1Y, tan2X, tan2Y, p2X, p2Y);
            codeLine = '';
            for (var k=0; k<curveCoordinates.length; k++){
                if (myCommand == 'c'){
                    trimX = lastX;
                    trimY = lastY;
                } else {
                    trimX = 0;
                    trimY = 0;
                }
                myX = curveCoordinates[k][0];
                myY = curveCoordinates[k][1];
                codeLine += getMoveToGCode(myX,myY,trimX,trimY,feedStr);
                lastX = myX;
                lastY = myY;
            }
            gcode += codeLine;
            lastX = myX;
            lastY= myY;
        }
    }

    return gcode;
}

function getMoveToGCode(myX,myY,trimX,trimY,feedStr){
    var gcode = '';
    var toX = roundDecimal((trimX + myX) * svgRatio);
    var toY = roundDecimal(TRIMORIGINY - (trimY + myY) * svgRatio);
    
    gcode += 'X' + toX + 'Y' + toY + feedStr + ' \n';
    
    return gcode;
}

function getHMoveToGCode(myX,trimX,feedStr){
    var gcode = '';
    var toX = roundDecimal((trimX + myX) * svgRatio);
    
    gcode += 'X' + toX + feedStr + ' \n';
    
    return gcode;
}

function getVMoveToGCode(myY,trimY,feedStr){
    var gcode = '';
    var toY = roundDecimal(TRIMORIGINY - (trimY + myY) * svgRatio);
    
    gcode += 'Y' + toY + feedStr + ' \n';
    
    return gcode;
}

////////////// BEZIER CURVE FUNCTIONS
function b3p0(t, p) {
  var k = 1 - t;
  return k * k * k * p;
}
function b3p1(t, p) {
  var k = 1 - t;
  return 3 * k * k * t * p;
}
function b3p2(t, p) {
  var k = 1 - t;
  return 3 * k * t * t * p;
}
function b3p3(t, p) {
  return t * t * t * p;
}
// Compute value on a single dimension for bezier curve with given control points
function b3(t, p0, p1, p2, p3) {
  return b3p0(t, p0) + b3p1(t, p1) + b3p2(t, p2) + b3p3(t, p3);
}
function createCurveCoordinates(subdivisions, p1X, p1Y, tan1X, tan1Y, tan2X, tan2Y, p2X, p2Y){
    var step;
    var myX;
    var myY;
    var coordinatesArr = new Array();
    
    for (var i=1; i<subdivisions; i++){
        step = i / subdivisions;
        
        myX = b3(step, p1X, tan1X, tan2X, p2X);
        myY = b3(step, p1Y, tan1Y, tan2Y, p2Y);
        
        coordinatesArr.push(new Array(myX, myY));
    }

    return coordinatesArr;
}