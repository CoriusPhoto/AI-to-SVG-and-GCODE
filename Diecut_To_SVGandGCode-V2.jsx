//*******************************************************
// Diecut_To_SVGandGCode.jsx
// Version 2.0
//
// Copyright 2026 Corius
// Comments or suggestions to contact@corius.fr
//
//*******************************************************

// Script version
var Diecut_To_SVGandGCode = 'v2.0';

// CONSOLE BLANK LINE
//$.writeln('          ----------<<<<<<<<<<     EXEC START     >>>>>>>>>>----------');

// AI document variables
var docObj = app.activeDocument;
var docName = docObj.fullName;
var docFolder = docName.parent.fsName;

// AI check variables
var OKtoCONTINUE = false;

// Data array variables
var AllPathsOfAllLayers = new Array(); // AllPathsOfAllLayers = [LayerPaths[], LayerPaths[], LayerPaths[],…]  //var allPathsByLayer = new Array();  // allPathsByLayer[i][LayerName, allPaths[]]
                                                            // LayerPaths = [PathData[], PathData[], PathData[],…]
                                                            // PathData = [LayerSettings, PATH, DIRECTION]
var AllPathsOfAllMergedLayers = new Array(); // AllPathsOfAllMergedLayers = [LayerPaths[], LayerPaths[], LayerPaths[],…] 
var AllPathsOfAllLayersMixedJobs = new Array(); // AllPathsOfAllLayersMixedJobs = [LayerPaths[], LayerPaths[], LayerPaths[],…] 
var AllPathsOrderedByLayer = new Array(); // AllPathsOrderedByLayer = [LayerPaths[], LayerPaths[], LayerPaths[],…]  //var allOrderedPathsByLayer = new Array();
var AllPathsOrderedMergedLayer = new Array(); // AllPathsOrderedMergedLayer = [LayerPaths[], LayerPaths[], LayerPaths[],…]  //var allOrderedPathsByLayer = new Array();
var AllPathsOrderedMixedJobs = new Array(); // AllPathsOrderedMixedJobs = [LayerPaths[], LayerPaths[], LayerPaths[],…]  //var allConvertedPathByLayer = new Array();
var MaterialProfile = new Array();  // MaterialProfile = [NAME, LayerSettings[], LayerSettings[], LayerSettings[], …]
                                                            // LayerSettings = [NAME, ORDER, SVGStoke[], PreviewStroke[], DashConvert[], MERGEABLE, POWER, SPEED, CYCLES]
                                                            // SVGStroke and PreviewStroke = [COLOR, THICKNESS]
                                                            // DashConvert = [PLAIN, GAP, POWER, SPEED

var ORIGIN = new Point();
ORIGIN.x = docObj.artboards[0].artboardRect[0];
ORIGIN.y = docObj.artboards[0].artboardRect[3];

var unitTable = new Array();
unitTable[0] = new Array('Full text', 'Unknown', 'Inches', 'Centimeters', 'Points', 'Picas', 'Millimeters', 'Qs', 'Pixels', 'FeetInches', 'Meters', 'Yards', 'Feet');
unitTable[1] = new Array('mm equivalent', 'Unknown', 25.4, 10, 0.3528, null, 1, null, 25.4/72, null, 1000, 914.4, 304.8);

// SVG variables
var unitRatio = null;
var svgRatio = 25.4/72;
var mmPointRatio = 0.3528;
var GLOBALSVG = true; // set to true to use global coordinates for SVG conversion, set to false to use local coodinates (local coordinates not yet implemented). 
var C = GLOBALSVG ? ' C' : ' c';
var S = GLOBALSVG ? ' S' : ' s';
var A = GLOBALSVG ? ' A' : ' a';
var L = GLOBALSVG ? ' L' : ' l';
var V = GLOBALSVG ? ' V' : ' v';
var H = GLOBALSVG ? ' H' : ' h';
var PRECISION = 1000; // the coordinates precision, rounded to 1/PRECISIONth (so 1000 means up to 3 digits after the dot, so 1/1000th)

// LASER GRBL VARIABLES
//var PARAMS = ''; /////////////// A SUPPRIMER, REMPLACER PAR LE NOM DU PROFILE DANS MATERIALPROFILE
var POWERMODE = 'M3'; // M3 for constant power, M4 for dynamic power
var TRIMORIGINY = 0;

var BEZIERSTEPS = 100;

// Preview variables
var MAKEPREVIEW = false;
var TRAVELSPEED = 1000;
var TRAVELSTROKE;
var LASTPREVIEWPOINT = new Point(0,0);
var ACCELERATION = 500;


// File variables
var MATERIALNAME = '';
var MixedJobsNames = new Array();
var SVGExportIndividual = false;
//var SVGExportMerged = false;
var SVGFullData = new Array();
var GCODEExportIndividual = false;
var GCODEExportMerged = false;
var GCODEExportMixed = false;
var GCODEFullData = new Array();
var INDIVIDUAL = '(indiv)';
var MERGED = '(merged)';
var MIXED = '(mixed)';

var svgFolder = '\\SVG';
var gcodeFolder = '\\GCODE';
var ExportedFiles = '';

checkMaterialSelected();

if (OKtoCONTINUE){
    getCommonSettings();
    getExportSettings();
    mainInitialize();
    getAllPaths();
    if (GCODEExportIndividual || SVGExportIndividual){
        optimizeTravel();
    }
    if (GCODEExportMerged && !GCODEExportMixed){
        optimizeMergedTravel();
    }
    if (GCODEExportMerged && GCODEExportMixed){
        optimizeMixedTravel();
    }
    convertPaths();
    makeGCODEfile();
    getAllSVGPaths();
    ExportAllFiles();
    if (MAKEPREVIEW){
        createPreview();
    }
    alert('Files Exported : '+ExportedFiles);
}


function checkMaterialSelected(){
    var myLayer = docObj.layers.getByName('GCODE_PARAMS');
    var SelectedProfileCount = 0;
    var myGroup;
    var myMaterialGroup;
    
    for (var i=0; i< myLayer.groupItems.length; i++){
         myGroup = myLayer.groupItems[i];
         if (myGroup.name != 'COMMON' && myGroup.hidden == false){
             SelectedProfileCount++;
             myMaterialGroup = myGroup;
         }
    }    
    
    if (SelectedProfileCount == 0){
        alert('Please select a material profile');
    } else if (SelectedProfileCount > 1){
        alert('Please select only 1 material profile');
    } else {
        OKtoCONTINUE = true;
        MATERIALNAME = myMaterialGroup.textFrames.getByName('MATERIAL_NAME').contents;
        getMaterialSettings(myMaterialGroup);
    }
}

function getMaterialSettings(myMaterialGroup){
    var mySettingsGroup = myMaterialGroup.groupItems.getByName('SETTINGS');
    var myGroup;
    var mySVGStrokeData = new Array();
    var myPreviewStroke;
    var myDashConvertData = new Array();
    var myLayerSettings = myMaterialGroup.groupItems.getByName('SETTINGS');
    
    var myOrder;
    var myMergeable;
    var myPower;
    var mySpeed;
    var myCycles;
    
    for (var i=0; i< mySettingsGroup.groupItems.length; i++){
        mySVGStrokeData = new Array();
        myPreviewStrokeData = new Array();
        myDashConvertData = new Array();
        myGroup = mySettingsGroup.groupItems[i];
        myName = myGroup.textFrames.getByName('JOB_NAME').contents;
        myOrder = myGroup.textFrames.getByName('JOB_ORDER').contents;
        myMergeable = (myGroup.textFrames.getByName('MERGEABLE').contents == 'YES')? true : false;
        myPower = 10 * myGroup.textFrames.getByName('POWER').contents;
        mySpeed = myGroup.textFrames.getByName('SPEED').contents;
        myCycles = myGroup.textFrames.getByName('CYCLES').contents;
        mySVGStrokeData[0] = myGroup.pathItems.getByName('SVG_PATH').strokeColor;
        mySVGStrokeData[1] = myGroup.pathItems.getByName('SVG_PATH').strokeWidth;
        myPreviewStroke = myGroup.pathItems.getByName('PREVIEW_PATH');
        myDashConvertData[0] = myGroup.textFrames.getByName('CONVERT_PLAIN').contents;
        myDashConvertData[1] = myGroup.textFrames.getByName('CONVERT_GAP').contents;
        myDashConvertData[2] = 10 * myGroup.textFrames.getByName('CONVERT_POWER').contents;
        myDashConvertData[3] = myGroup.textFrames.getByName('CONVERT_SPEED').contents;
        myLayerSettings = new Array();
        myLayerSettings[0] = myName;
        myLayerSettings[1] = myOrder;
        myLayerSettings[2] = mySVGStrokeData;
        myLayerSettings[3] = myPreviewStroke;
        myLayerSettings[4] = myDashConvertData;
        myLayerSettings[5] = myMergeable;
        myLayerSettings[6] = myPower;
        myLayerSettings[7] = mySpeed;
        myLayerSettings[8] = myCycles;
        
        MaterialProfile.push(myLayerSettings); 
    }
    MaterialProfile.sort(triAscendant);
} 

function triAscendant(a, b) {
    return a[1] - b[1];
}

function getCommonSettings(){
    var myLayer = docObj.layers.getByName('GCODE_PARAMS');
    var myGroup = myLayer.groupItems.getByName('COMMON').groupItems.getByName('TOP_SETTINGS');
    
    PRECISION = myGroup.groupItems.getByName('CALCUL').textFrames.getByName('PRECISION').contents;    
    BEZIERSTEPS = myGroup.groupItems.getByName('CALCUL').textFrames.getByName('BEZIER').contents;
    TRAVELSPEED = myGroup.groupItems.getByName('CALCUL').textFrames.getByName('OFF_SPEED').contents;
    TRAVELSTROKE = myGroup.groupItems.getByName('CALCUL').pathItems.getByName('TRAVEL_PATH');
    var YesNo = myGroup.groupItems.getByName('CALCUL').textFrames.getByName('MAKE_PREVIEW').contents;
    MAKEPREVIEW = (YesNo == 'YES')? true : false;
    ACCELERATION = myGroup.groupItems.getByName('CALCUL').textFrames.getByName('ACCEL_FACTOR').contents;
}

function getExportSettings(){
    var myLayer = docObj.layers.getByName('GCODE_PARAMS');
    var myGroup = myLayer.groupItems.getByName('COMMON').groupItems.getByName('TOP_SETTINGS');
    
    var YesNo = myGroup.groupItems.getByName('EXPORT').textFrames.getByName('SVG_INDIV').contents;
    SVGExportIndividual = (YesNo == 'YES')? true : false;
    //YesNo = myGroup.groupItems.getByName('EXPORT').textFrames.getByName('SVG_MERGE').contents;
    //SVGExportMerged = (YesNo == 'YES')? true : false;
    YesNo = myGroup.groupItems.getByName('EXPORT').textFrames.getByName('GCODE_INDIV').contents;
    GCODEExportIndividual = (YesNo == 'YES')? true : false;
    YesNo = myGroup.groupItems.getByName('EXPORT').textFrames.getByName('GCODE_MERGE').contents;
    GCODEExportMerged = (YesNo == 'YES')? true : false;
    YesNo = myGroup.groupItems.getByName('EXPORT').textFrames.getByName('GCODE_MIX').contents;
    GCODEExportMixed = (YesNo == 'YES')? true : false;
}

function mainInitialize(){
    var w = Math.round(100 * docObj.width)/100;
    var h = Math.round(100 * docObj.height)/100;
    
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
        }
    }

    TRIMORIGINY = Math.round(PRECISION * h * svgRatio)/PRECISION;
} 

function getAllPaths() {
    var myLayer;
    var myPath;
    var myPathData;
    var myLayerName;
    var myLayerSettings;
    var myLayerPaths;
    var myMergeLayerPaths = new Array();
    var pathOK;
    var myMergeable;
    var myMergeCycles;
    var merging = false;
    var mergingCycles = 0;
    var myDashedPaths;
    var tempoLayerSettings;
    var pathInGroupsArr;
    
    for (var i=0;i<MaterialProfile.length;i++){
        myLayerSettings = MaterialProfile[i];
        myLayerName = myLayerSettings[0];
        
        try {myLayer = docObj.layers.getByName(myLayerName);} catch (e) {}
        
        if (myLayer != null){            
            myLayerPaths = new Array();
            
            for (var j=0;j<myLayer.pathItems.length && myLayer.visible == true;j++){
                myPath = myLayer.pathItems[j];
                pathOK = checkPathInArtboard(myPath);
                if (myPath.hidden == false && pathOK && myPath.pathPoints.length > 1 ){
                    if (myLayerSettings[4][0] > 0 && myLayerSettings[4][1] > 0){
                        myDashedPaths = dashedPath(myPath, myLayerSettings[4][0]/mmPointRatio, myLayerSettings[4][1]/mmPointRatio);
                        tempoLayerSettings = copyArray (myLayerSettings);
                        tempoLayerSettings[6] = myLayerSettings[4][2];
                        tempoLayerSettings[7] = myLayerSettings[4][3];
                    } else {
                        myDashedPaths = new Array(myPath);
                        tempoLayerSettings = myLayerSettings;
                    }
                    for (var k=0;k<myDashedPaths.length;k++){
                        myPathData = new Array();
                        myPathData[0] = (myDashedPaths.length == 1)? myLayerSettings : tempoLayerSettings;
                        myPathData[1] = myDashedPaths[k];
                        myPathData[2] = 1; // This is the PathDirection, will be updated to -1 if path should be reversed for travel optimization
                        
                        myLayerPaths.push(myPathData);
                    }
                }
            }
        
            for (j=0;j<myLayer.groupItems.length && myLayer.visible == true;j++){
                myGroup = myLayer.groupItems[j];
                pathInGroupsArr = getPathsInGroup(myGroup);
                for (j=0;j<pathInGroupsArr.length;j++){
                    myPath = pathInGroupsArr[j];
                    pathOK = checkPathInArtboard(myPath);
                    if (myPath.hidden == false && pathOK){
                        if (myLayerSettings[4][0] > 0 && myLayerSettings[4][1] > 0){
                            myDashedPaths = dashedPath(myPath, myLayerSettings[4][0]/mmPointRatio, myLayerSettings[4][1]/mmPointRatio);
                            tempoLayerSettings = copyArray (myLayerSettings);
                            tempoLayerSettings[6] = myLayerSettings[4][2];
                            tempoLayerSettings[7] = myLayerSettings[4][3];
                        } else {
                            myDashedPaths = new Array(myPath);
                            tempoLayerSettings = myLayerSettings;
                        }
                        for (var k=0;k<myDashedPaths.length;k++){
                            myPathData = new Array();
                            myPathData[0] = (myDashedPaths.length == 1)? myLayerSettings : tempoLayerSettings;
                            myPathData[1] = myDashedPaths[k];
                            myPathData[2] = 1; // This is the PathDirection, will be updated to -1 if path should be reversed for travel optimization
                            
                            myLayerPaths.push(myPathData);
                        }
                    }
                }
            }
            if (myLayerPaths.length > 0){
                AllPathsOfAllLayers.push(myLayerPaths);
            } 
        }
    }
    AllPathsOfAllMergedLayers=copyArray(AllPathsOfAllLayers);
    AllPathsOfAllLayersMixedJobs=creerNouveauTableauFusionne(copyArray(AllPathsOfAllLayers));
}

function getPathsInGroup(myGroup){
    var myPathArr = new Array();
    var tempArr;
    
    for (var i=0 ; i<myGroup.pathItems.length ; i++){
        myPathArr[myPathArr.length] = myGroup.pathItems[i];
    }
    
    for (var i=0 ; i<myGroup.groupItems.length ; i++){
        tempArr = getPathsInGroup(myGroup.groupItems[i]);
        for (var j=0;j<tempArr.length;j++){
            myPathArr[myPathArr.length] = tempArr[j];
        }
    }

    return myPathArr;
}

function copyArray(myArray) {
    var copie = [];
    for (var i = 0; i < myArray.length; i++) {
        if (myArray[i] instanceof Array) {
            copie[i] = copyArray(myArray[i]); // Récursif pour les sous-tableaux
        } else {
            copie[i] = myArray[i];
        }
    }
    return copie;
}

function creerNouveauTableauFusionne(myArray) {
    var resultatPaths = [];
    var i = 0;
    var j;
    var fusion;
    
    while (i < myArray.length) {
        var pathsFusionnes = myArray[i].slice(); // Copie du premier tableau
        MixedJobsNames[i] = myArray[i][0][0][0];
        // Tant que le suivant a les mêmes CYCLES, on fusionne
        fusion = false;
        j = 1;
        while (i + j < myArray.length && 
           i + j < MaterialProfile.length &&
           MaterialProfile[i][5] &&
           MaterialProfile[i][8] === MaterialProfile[i + j][8]) {
                
            MixedJobsNames[i] += '~'+myArray[i+j][0][0][0];
            pathsFusionnes = pathsFusionnes.concat(myArray[i + j]);
            j++; // On avance dans la boucle interne
            fusion = true;
        }
        
        resultatPaths.push(pathsFusionnes);
        //resultatProfiles.push(MaterialProfile[i]); // Garder le profil correspondant
        
        i+=(fusion)? j+1 : 1; // Passer au groupe suivant
    }

    return resultatPaths;
}

function tableauxEgaux(tab1, tab2) {
    var isSame = false;
    // Même référence ?
    if (tab1 === tab2){
        isSame = true;
    }
    
    // Vérifications de base
    if (!tab1 || !tab2){
        isSame = false;
    }
    if (tab1.length !== tab2.length){
        isSame = false;
    }
    
    // Comparer chaque élément
    for (var i = 0; i < tab1.length; i++) {
        var elem1 = tab1[i];
        var elem2 = tab2[i];
        
        // Si les éléments sont des tableaux, comparaison récursive
        if (elem1 instanceof Array && elem2 instanceof Array) {
            if (!tableauxEgaux(elem1, elem2)) {
                isSame = false;
            }
        }
        // Sinon, comparaison simple
        else {
            isSame = (elem1 !== elem2)? false : true;
        }
    }
    
    return isSame;
}

function checkPathInArtboard(myPath){
    var inArtBoard = false;
    
    if (myPath.geometricBounds[0] >= 0 && myPath.geometricBounds[1] <= 0 && myPath.geometricBounds[2] <= docObj.artboards[0].artboardRect[2] && myPath.geometricBounds[3] >= docObj.artboards[0].artboardRect[3]){
        inArtBoard = true;
    }
        
    return inArtBoard;
}

function optimizeTravel(){
    var myPathData;
    var myPath;
    var bestIndex;
    var bestDist;
    var bestDirection;
    var bestEndPoint;
    var currentPoint = new Point();
    var myStart;
    var myEnd;
    var distance1;
    var distance2;
    
    currentPoint.x = ORIGIN.x;
    currentPoint.y = ORIGIN.y;
    for (var i=0;i<AllPathsOfAllLayers.length;i++){
        AllPathsOrderedByLayer[i] = new Array();
        currentPoint.x = ORIGIN.x;
        currentPoint.y = ORIGIN.y;
        
        while (AllPathsOfAllLayers[i].length > 0){
            bestIndex = 0;
            bestDist = Infinity;
            distance2 = Infinity;
            bestDirection = 1;
            bestEndPoint;
            
            for (var j=0;j<AllPathsOfAllLayers[i].length;j++){
                myPathData = AllPathsOfAllLayers[i][j];
                myPath = myPathData[1];
                
                myStart = getStartPoint(myPath);
                myEnd = getEndPoint(myPath);
                
                distance1 = euclideanDistance(currentPoint, myStart);
                
                if (distance1 < bestDist){
                    bestDist = distance1;
                    bestIndex = j;
                    bestDirection = 1;
                    bestEndPoint = myEnd;
                    bestPath = myPath;
                    bestPathStart = myStart;
                    bestPathEnd = myEnd;
                }
                
                if (!myPath.closed){
                    distance2 = euclideanDistance(currentPoint, myEnd);
                    if (distance2 < bestDist){
                        bestDist = distance2;
                        bestIndex = j;
                        bestDirection = -1;
                        bestEndPoint = myStart;
                        bestPath = myPath;
                        bestPathStart = myStart;
                        bestPathEnd = myEnd;
                    }
                }
            }
        
            currentPoint = bestEndPoint;
            AllPathsOrderedByLayer[i].push(AllPathsOfAllLayers[i][bestIndex]);
            AllPathsOrderedByLayer[i][AllPathsOrderedByLayer[i].length - 1][2] = bestDirection;
            AllPathsOfAllLayers[i].splice(bestIndex,1);
        }
    }
}

function optimizeMergedTravel(){
    var myPathData;
    var myPath;
    var bestIndex;
    var bestDist;
    var bestDirection;
    var bestEndPoint;
    var currentPoint = new Point();
    var myStart;
    var myEnd;
    var distance1;
    var distance2;
    var newMergeable = false;
    var merging = false;
    
    currentPoint.x = ORIGIN.x;
    currentPoint.y = ORIGIN.y;
    
    for (var i=0;i<AllPathsOfAllMergedLayers.length;i++){
        AllPathsOrderedMergedLayer[i] = new Array();
        newMergeable = MaterialProfile[i][5];
        if (!merging){
            merging = newMergeable;
            currentPoint.x = ORIGIN.x;
            currentPoint.y = ORIGIN.y;
        } else if (!newMergeable){
            merging = newMergeable;
            currentPoint.x = ORIGIN.x;
            currentPoint.y = ORIGIN.y;
        }
        while (AllPathsOfAllMergedLayers[i].length > 0){
            bestIndex = 0;
            bestDist = Infinity;
            distance2 = Infinity;
            bestDirection = 1;
            bestEndPoint = currentPoint;
            
            for (var j=0;j<AllPathsOfAllMergedLayers[i].length;j++){
                myPathData = AllPathsOfAllMergedLayers[i][j];
                myPath = myPathData[1];
                
                myStart = getStartPoint(myPath);
                myEnd = getEndPoint(myPath);
                
                distance1 = euclideanDistance(currentPoint, myStart);
                
                if (distance1 < bestDist){
                    bestDist = distance1;
                    bestIndex = j;
                    bestDirection = 1;
                    bestEndPoint = myEnd;
                }
                
                if (!myPath.closed){
                    distance2 = euclideanDistance(currentPoint, myEnd);
                    if (distance2 < bestDist){
                        bestDist = distance2;
                        bestIndex = j;
                        bestDirection = -1;
                        bestEndPoint = myStart;
                    }
                }
            }
            
            currentPoint = bestEndPoint;
            AllPathsOrderedMergedLayer[i].push(AllPathsOfAllMergedLayers[i][bestIndex]);
            AllPathsOrderedMergedLayer[i][AllPathsOrderedMergedLayer[i].length - 1][2] = bestDirection;
            AllPathsOfAllMergedLayers[i].splice(bestIndex,1);
        }
    }
}

function optimizeMixedTravel(){
    var myPathData;
    var myPath;
    var bestIndex;
    var bestDist;
    var bestDirection;
    var bestEndPoint;
    var currentPoint = new Point();
    var myStart;
    var myEnd;
    var distance1;
    var distance2;
    
    currentPoint.x = ORIGIN.x;
    currentPoint.y = ORIGIN.y;
     
    for (var i=0;i<AllPathsOfAllLayersMixedJobs.length;i++){
        AllPathsOrderedMixedJobs[i] = new Array();
        currentPoint.x = ORIGIN.x;
        currentPoint.y = ORIGIN.y;
        while (AllPathsOfAllLayersMixedJobs[i].length > 0){
            bestIndex = 0;
            bestDist = Infinity;
            distance2 = Infinity;
            bestDirection = 1;
            bestEndPoint;
            
            for (var j=0;j<AllPathsOfAllLayersMixedJobs[i].length;j++){
                myPathData = AllPathsOfAllLayersMixedJobs[i][j];
                myPath = myPathData[1];
                
                myStart = getStartPoint(myPath);
                myEnd = getEndPoint(myPath);
                
                distance1 = euclideanDistance(currentPoint, myStart);
                
                if (distance1 < bestDist){
                    bestDist = distance1;
                    bestIndex = j;
                    bestDirection = 1;
                    bestEndPoint = myEnd;
                }
                
                if (!myPath.closed){
                    distance2 = euclideanDistance(currentPoint, myEnd);
                    if (distance2 < bestDist){
                        bestDist = distance2;
                        bestIndex = j;
                        bestDirection = -1;
                        bestEndPoint = myStart;
                    }
                }
            }
            AllPathsOrderedMixedJobs[i].push(AllPathsOfAllLayersMixedJobs[i][bestIndex]);
            AllPathsOrderedMixedJobs[i][AllPathsOrderedMixedJobs[i].length - 1][2] = bestDirection;
            
            currentPoint = bestEndPoint;
            AllPathsOfAllLayersMixedJobs[i].splice(bestIndex,1);
        }
    }
}

function euclideanDistance(p1,p2){
    var dx = p1.x - p2.x;
    var dy = p1.y - p2.y;
    return Math.sqrt(dx*dx + dy*dy);
}


function getStartPoint(myPath){

    var pt = new Point();

    if(myPath.pathPoints){  // PathItem Illustrator

        pt.x = myPath.pathPoints[0].anchor[0];
        pt.y = myPath.pathPoints[0].anchor[1];

    }else if(myPath.length){  // tableau de points

        pt.x = myPath[0].anchor[0];
        pt.y = myPath[0].anchor[1];

    }else{

        throw new Error("getStartPoint : unsupported object");
    }

    return pt;
}

function getEndPoint(myPath){
    var last;

    var pt = new Point();

    if(myPath.pathPoints){  // PathItem Illustrator
        last = myPath.pathPoints.length - 1;
        pt.x = myPath.pathPoints[last].anchor[0];
        pt.y = myPath.pathPoints[last].anchor[1];

    }else if(myPath.length){  // tableau de points
        last = myPath.length - 1;
        pt.x = myPath[last].anchor[0];
        pt.y = myPath[last].anchor[1];

    }else{

        throw new Error("getEndPoint : unsupported object");
    }

    return pt;
}

function convertPaths(){
    var myLayerPaths;
    var myPathData;
    var i;
    var j;
    

    if(GCODEExportIndividual){
        //convert path in AllPathsOrderedByLayer into GCODE paths
        for (i=0;i<AllPathsOrderedByLayer.length;i++){
            myLayerPaths = AllPathsOrderedByLayer[i];
            for (j=0;j<myLayerPaths.length;j++){
                myPathData = myLayerPaths[j];
                convertIllustratorPathToGCODE(myPathData);
            }
        }
    }
    if(GCODEExportMerged && !GCODEExportMixed){
        //convert path in AllPathsOrderedByLayer into GCODE paths
        for (i=0;i<AllPathsOrderedMergedLayer.length;i++){
            myLayerPaths = AllPathsOrderedMergedLayer[i];
            for (j=0;j<myLayerPaths.length;j++){
                myPathData = myLayerPaths[j];
                convertIllustratorPathToGCODE(myPathData);
            }
        }
    }
    if(GCODEExportMerged && GCODEExportMixed){
        //convert path in AllPathsOrderedMixedJobs into GCODE paths
        for (i=0;i<AllPathsOrderedMixedJobs.length;i++){
            myLayerPaths = AllPathsOrderedMixedJobs[i];
            for (j=0;j<myLayerPaths.length;j++){
                myPathData = myLayerPaths[j];
                convertIllustratorPathToGCODE(myPathData);
            }
        }
    }
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

///////////////////////////////////////////////
///////////////////////////////////////////////
function convertIllustratorPathToGCODE(PathData){

    var layerSettings = PathData[0];
    var myPath = PathData[1];
    var myDirection = PathData[2];

    var laserPower = layerSettings[6];
    var laserSpeed = layerSettings[7];

    var gcode = "";

    var totalLoop;
    var startPathPointID;
    var sameArrayContent = false;
    
    if(myPath.pathPoints){  // PathItem Illustrator
        totalLoop = myPath.closed ? myPath.pathPoints.length + 1 : myPath.pathPoints.length;
        startPathPointID = (myDirection > 0 || myPath.closed) ? 0 : myPath.pathPoints.length - 1;
    }else if(myPath.length){  // tableau de points
        sameArrayContent = tableauxEgaux(myPath[0], myPath[myPath.length - 1]);
        totalLoop = sameArrayContent ? myPath.length + 1 : myPath.length;
        startPathPointID = (myDirection > 0 || sameArrayContent) ? 0 : myPath.length - 1;
    }

    var mySegPoint1;
    var mySegPoint2;

    var P1, P2;
    var anteP1, postP1;
    var anteP2, postP2;
    
    var myPointID;
    var myPointID2;

    for (var loopCount = 0 ; loopCount < totalLoop-1 ; loopCount++){
        if(myPath.pathPoints){
            myPointID = getID(myPath.pathPoints,startPathPointID + loopCount * myDirection);
            myPointID2 = getID(myPath.pathPoints,myPointID + myDirection);

            mySegPoint1 = myPath.pathPoints[myPointID];
            mySegPoint2 = myPath.pathPoints[myPointID2];
        }else if(myPath.length){
            myPointID = getID(myPath,startPathPointID + loopCount * myDirection);
            myPointID2 = getID(myPath,myPointID + myDirection);

            mySegPoint1 = myPath[myPointID];
            mySegPoint2 = myPath[myPointID2];
        }

        P1 = mySegPoint1.anchor;
        P2 = mySegPoint2.anchor;

        anteP1 = (myDirection > 0) ? mySegPoint1.leftDirection  : mySegPoint1.rightDirection;
        postP1 = (myDirection > 0) ? mySegPoint1.rightDirection : mySegPoint1.leftDirection;

        anteP2 = (myDirection > 0) ? mySegPoint2.leftDirection  : mySegPoint2.rightDirection;
        postP2 = (myDirection > 0) ? mySegPoint2.rightDirection : mySegPoint2.leftDirection;

        if(loopCount == 0){
            gcode += '\nG0 M5';
            gcode += '\nX'+toMMX(P1[0])+' Y'+toMMY(P1[1]);
            gcode += '\n'+POWERMODE+' S'+laserPower+' G1 F'+laserSpeed;
        }

        var isLine = (P1[0]==postP1[0] && P1[1]==postP1[1]) &&(P2[0]==anteP2[0] && P2[1]==anteP2[1]);

        if(isLine){
            gcode += '\nX'+toMMX(P2[0])+' Y'+toMMY(P2[1]);
        }
        else{

            for(var s=1;s<=BEZIERSTEPS;s++){

                var t = s/BEZIERSTEPS;

                var pt = bezier(P1,postP1,anteP2,P2,t);
                gcode += '\nX'+toMMX(pt[0])+ ' Y'+toMMY(pt[1]);
            }
        }
    }

    PathData[4] = gcode;
}


function roundCoord(v){
    return Math.round(v * PRECISION) / PRECISION;
}

function toMMX(x){
    return roundCoord((x - ORIGIN.x) * svgRatio);
}

function toMMY(y){
    return roundCoord((y - ORIGIN.y) * svgRatio);
}

function getID(pathpoints,id){

    if (id < 0){
        return pathpoints.length + id;
    }
    else if (id > pathpoints.length - 1){
        return id - pathpoints.length;
    }
    else{
        return id;
    }
}

function bezier(A,B,C,D,t){

    var u = 1 - t;

    var x =
        u*u*u*A[0] +
        3*u*u*t*B[0] +
        3*u*t*t*C[0] +
        t*t*t*D[0];

    var y =
        u*u*u*A[1] +
        3*u*u*t*B[1] +
        3*u*t*t*C[1] +
        t*t*t*D[1];

    return [x,y];
}

//////////////////////////////////////////////////////////////////////////// COLOR PROCESSING FUNCTIONS
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
//////////////////////////////////////////////////////////////

function getAllSVGPaths(){
    for (var i=0; i<AllPathsOrderedByLayer.length; i++){
        if (AllPathsOrderedByLayer[i][0][1] != null && AllPathsOrderedByLayer[i][0][1].length >0){
            getSVGPaths(i);
        }
     }
    makeSVGfile(i);
}

function getSVGPaths(LayerID){
    var myLayerPath = AllPathsOrderedByLayer[LayerID];
    var svgTXT = '';
    var layerName =  AllPathsOrderedByLayer[LayerID][0][0][0];
    
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
    var myPointID2;
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
    var PathLENGTH;
    
    //$.writeln('myPathsArr.length : '+myPathsArr.length);
    for (pathID=0; pathID<myLayerPath.length;pathID++){
        svgTXT = '<path class="'+layerName+'" d="M';
        myPath = myLayerPath[pathID][1];
        myDirection = myLayerPath[pathID][2];
        previousCurved = false;
    
        if(myPath.pathPoints){  // PathItem Illustrator
            PathLENGTH = myPath.pathPoints.length;
            totalLoop = myPath.closed ? myPath.pathPoints.length + 1 : myPath.pathPoints.length;
            startPathPointID = (myDirection > 0 || myPath.closed) ? 0 : myPath.pathPoints.length - 1;
            sameArrayContent = myPath.closed;
        }else if(myPath.length){  // tableau de points
            PathLENGTH = myPath.length;
            totalLoop = myPath.closed ? myPath.length + 1 : myPath.length;
            sameArrayContent = tableauxEgaux(myPath[0], myPath[myPath.length - 1]);
            startPathPointID = (myDirection > 0 || sameArrayContent) ? 0 : myPath.length - 1;
        }
    
        for (loopCount = 0 ; loopCount < PathLENGTH ; loopCount++){
            if(myPath.pathPoints){
                myPointID = getID(myPath.pathPoints,startPathPointID + loopCount * myDirection);
                mySegPoint1 = myPath.pathPoints[myPointID];
            }else if(myPath.length){
                myPointID = getID(myPath,startPathPointID + loopCount * myDirection);
                mySegPoint1 = myPath[myPointID];
            }

            P1id = myPointID;
            if (loopCount < PathLENGTH - 1) {
                if(myPath.pathPoints){
                    myPointID2 = getID(myPath.pathPoints,myPointID + myDirection);
                    mySegPoint2 = myPath.pathPoints[myPointID2];
                }else if(myPath.length){
                    myPointID2 = getID(myPath,myPointID + myDirection);
                    mySegPoint2 = myPath[myPointID2];
                }

                P2id = myPointID + myDirection;
            } else {
                if(myPath.pathPoints){
                    myPointID2 =  (myDirection > 0) ? getID(myPath.pathPoints,0) : getID(myPath.pathPoints,myPath.pathPoints.length - 1);
                    mySegPoint2 = myPath.pathPoints[myPointID2];
                    P2id = (myDirection > 0) ? 0 : myPath.pathPoints.length - 1;
                }else if(myPath.length){
                    myPointID2 = (myDirection > 0) ? getID(myPath,0) : getID(myPath,myPath.length - 1);
                    mySegPoint2 = myPath[myPointID2];
                    P2id = (myDirection > 0) ? 0 : myPath.length - 1;
                }
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
        
            if (loopCount < PathLENGTH - 1){
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
            } else if(sameArrayContent){
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
        myLayerPath[pathID][3] = svgTXT;
    }
}

function makeSVGfile(){
    var myLayerPath;
    var myJobName;
    var svgstring;
    var fullsvgstring;
    var w = Math.round(100 * docObj.width)/100;
    var h = Math.round(100 * docObj.height)/100;
    var strokeColor;
    var strokeCMYKColor;
    var svgStart = '<?xml version="1.0" encoding="UTF-8"?>\n<svg  id="CoriusSVG-GCODE" data-name="';    
    var svgendheadertxt = 'xmlns="http://www.w3.org/2000/svg"';
    var svgEnd = '</svg>';
    var svgDefstring;
    var i;
    var j;
    var k;
    var myExt = '.svg';
    var myCycles;
    var merging = false;
    var newMergeable = false;
        
    svgendheadertxt += ' width="' +Math.round(PRECISION * w * svgRatio)/PRECISION+ 'mm" height="' +Math.round(PRECISION * h * svgRatio)/PRECISION+ 'mm" viewBox="0 0 ' +w+ ' ' +h+'">\n'; 
    
    if (SVGExportIndividual){
        for (i=0;i<AllPathsOrderedByLayer.length;i++){
            svgstring = '';
            myLayerPath = AllPathsOrderedByLayer[i];
            myJobName = myLayerPath[0][0][0];
            strokeCMYKColor = myLayerPath[0][0][2][0];
            strokeColor = cmykToHex (strokeCMYKColor.cyan, strokeCMYKColor.magenta, strokeCMYKColor.yellow, strokeCMYKColor.black);
            svgDefstring = '<defs>\n<style>\n.'+myJobName+' {\nfill: none;\nstroke: '+strokeColor+';\nstroke-miterlimit: 10;\nstroke-width: .5px;\n}\n</style>\n</defs>\n';
            for (j=0;j<myLayerPath.length;j++){
                svgstring += myLayerPath[j][3];
            }
            fullsvgstring = svgStart+myJobName+'" '+svgendheadertxt+ svgDefstring+ svgstring + svgEnd;
            k = SVGFullData.length;
            SVGFullData[k] = new Array();
            SVGFullData[k][1] = myExt;
            SVGFullData[k][2] = myJobName;
            SVGFullData[k][3] = INDIVIDUAL;
            SVGFullData[k][4] = fullsvgstring;
        }
    }
}

function makeGCODEfile(){
    var myLayerPath;
    var myJobName;
    var gcodestring;
    var gcodeCycleString;
    var fullgcodestring;
    var gcodeStart = 'G90';
    var gcodeEnd = '\nG0 M5\nX0 Y0 Z0';
    var i;
    var j;
    var k;
    var myExt = '.nc';
    var myCycles;
    var merging = false;
    var newMergeable = false;
    
    if (GCODEExportIndividual){
        for (i=0;i<AllPathsOrderedByLayer.length;i++){
            gcodestring = '';
            myLayerPath = AllPathsOrderedByLayer[i];
            myJobName = myLayerPath[0][0][0];
            myCycles = myLayerPath[0][0][8];
            for (j=0;j<myLayerPath.length;j++){
                gcodestring += myLayerPath[j][4];
            }
            for (j=1;j<myCycles;j++){
                gcodestring += gcodestring;
            }
            fullgcodestring = gcodeStart + gcodestring + gcodeEnd;
            k = GCODEFullData.length;
            GCODEFullData[k] = new Array();
            GCODEFullData[k][1] = myExt;
            GCODEFullData[k][2] = myJobName;
            GCODEFullData[k][3] = INDIVIDUAL;
            GCODEFullData[k][4] = fullgcodestring;
        }
    }
    
    if (GCODEExportMerged && !GCODEExportMixed){
        for (i=0;i<AllPathsOrderedMergedLayer.length;i++){
            newMergeable = AllPathsOrderedMergedLayer[i][0][0][5];
            if (!merging){
                merging = newMergeable;
                gcodestring = '';
                myJobName = '';
            } else if (!newMergeable){
                merging = newMergeable;
                gcodestring = '';
                myJobName = '';
            }
            gcodeCycleString = '';
            myLayerPath = AllPathsOrderedMergedLayer[i];
            myJobName += (myJobName == '') ? myLayerPath[0][0][0] : '+'+myLayerPath[0][0][0];
            myCycles = myLayerPath[0][0][8];
            for (j=0;j<myLayerPath.length;j++){
                gcodeCycleString += myLayerPath[j][4];
            }
            for (j=0;j<myCycles;j++){
                gcodestring += gcodeCycleString;
            }
            if (!merging || i == AllPathsOrderedMergedLayer.length-1 || (i < AllPathsOrderedMergedLayer.length - 1 && !AllPathsOrderedMergedLayer[i+1][0][0][5])){
                fullgcodestring = gcodeStart + gcodestring + gcodeEnd;
                k = GCODEFullData.length;
                GCODEFullData[k] = new Array();
                GCODEFullData[k][1] = myExt;
                GCODEFullData[k][2] = myJobName;
                GCODEFullData[k][3] = MERGED;
                GCODEFullData[k][4] = fullgcodestring;
            }
        }
    }
        
    gcodestring = '';
    
    if (GCODEExportMerged && GCODEExportMixed){
        for (i=0;i<AllPathsOrderedMixedJobs.length;i++){
            newMergeable = AllPathsOrderedMixedJobs[i][0][0][5];
            if (!merging){
                merging = newMergeable;
                gcodestring = '';
                myJobName = '';
            } else if (!newMergeable){
                merging = newMergeable;
                gcodestring = '';
                myJobName = '';
            }
            gcodeCycleString = '';
            myLayerPath = AllPathsOrderedMixedJobs[i];
            myJobName += (myJobName == '') ? myLayerPath[0][0][0] : '+'+myLayerPath[0][0][0];
            myCycles = myLayerPath[0][0][8];
            for (j=0;j<myLayerPath.length;j++){
                gcodeCycleString += myLayerPath[j][4];
            }
            for (j=0;j<myCycles;j++){
                gcodestring += gcodeCycleString;
            }
            if (!merging || i == AllPathsOrderedMixedJobs.length-1 || (i < AllPathsOrderedMixedJobs.length - 1 && !AllPathsOrderedMixedJobs[i+1][0][0][5])){
                fullgcodestring = gcodeStart + gcodestring + gcodeEnd;
                k = GCODEFullData.length;
                GCODEFullData[k] = new Array();
                GCODEFullData[k][1] = myExt;
                GCODEFullData[k][2] = myJobName;
                GCODEFullData[k][3] = MIXED;
                GCODEFullData[k][4] = fullgcodestring;
            }
        }
    }
}

function ExportAllFiles(){
    var myJobName;
    var myExt;
    var contentString;
    var i;
    
    for (i=0;i<GCODEFullData.length;i++){
        myJobName = GCODEFullData[i][2]+GCODEFullData[i][3];
        myExt = GCODEFullData[i][1];
        contentString = GCODEFullData[i][4];
        saveFile(myJobName,myExt,contentString);
    }

    for (i=0;i<SVGFullData.length;i++){
        //myJobName = SVGFullData[i][2]+SVGFullData[i][3];
        myJobName = SVGFullData[i][2];
        myExt = SVGFullData[i][1];
        contentString = SVGFullData[i][4];
        saveFile(myJobName,myExt,contentString);
    }
}

function saveFile(myJobName,myExt,contentString){
    var myFileName;
    var myFile;
    var myFolderName = (myExt == '.svg')? docFolder + svgFolder : docFolder + gcodeFolder;
    var myFolder = new Folder(myFolderName);
    
    myFileName = docObj.name.substring(0,docObj.name.lastIndexOf ('.ai'))+'_'+MATERIALNAME+'_'+myJobName+myExt;
    myFolder.create();
    
    myFile = new File(myFolderName+'\\'+myFileName);
    myFile.encoding = "BINARY";
    myFile.open('w');
    myFile.write(contentString);
    myFile.close();
    ExportedFiles += '\n'+myFileName;
    
}

///////////////////////// PREVIEW
function createPreview(){
    var myLayerName;
    var myJobData;
    
    myLayerName = "LASER_PREVIEW";
    cleanLayer(myLayerName);    
    if (GCODEExportIndividual){
        myJobData = AllPathsOrderedByLayer;
        generateLaserPreviewLayer(myLayerName,myJobData);
    }
    
    myLayerName = "Merged Files LASER_PREVIEW";
    cleanLayer(myLayerName);    
    if (GCODEExportMerged && !GCODEExportMixed){
        myJobData = AllPathsOrderedMergedLayer;
        generateLaserPreviewLayer(myLayerName,myJobData);
    }
    
    myLayerName = "Mixed Jobs LASER_PREVIEW";
    cleanLayer(myLayerName);    
    if (GCODEExportMerged && GCODEExportMixed){
        myJobData = AllPathsOrderedMixedJobs;
        generateLaserPreviewLayer(myLayerName,myJobData);
    }
}

function cleanLayer(myLayerName){
    var myLayer;
    
    for (var i=0;i<docObj.layers.length;i++){
        myLayer = docObj.layers[i];
        if (myLayer.name == myLayerName){
            myLayer.locked = false;
            myLayer.visible = true;
            myLayer.remove();
            i--;
        }        
    }
}

function generateLaserPreviewLayer(myLayerName,myJobData) {    
    var myLayerPath;
    var myPreviewLayer = docObj.layers.add();
    var elapsedJobTime = 0;
    var elapsedTotalTime = 0;
    var timeText = '';
    var seconds;
    var merging = false;
    var Mergeable = false;
    var newMergeable = false;
    myPreviewLayer.name = myLayerName;
    myPreviewLayer.zOrder(ZOrderMethod.SENDTOBACK);
    
    LASTPREVIEWPOINT.x = 0;
    LASTPREVIEWPOINT.y = 0;
    
    var backToZeroAfterJob = (myLayerName != 'LASER_PREVIEW')? false : true;
    
    for (var i=0;i<myJobData.length;i++){
        timeText = 'Job : ';
        myLayerPath = myJobData[i];
        if (myLayerPath[0][0][3].stroked){
            if(myLayerName != 'LASER_PREVIEW'){
                if (i<myJobData.length - 1){
                    Mergeable = myJobData[i][0][0][5];
                    newMergeable = myJobData[i+1][0][0][5];
                    if (Mergeable && newMergeable){
                        backToZeroAfterJob = false;
                    } else {
                        backToZeroAfterJob = true;
                    }
                } else {
                    backToZeroAfterJob = true;
                }
                //$.writeln('>>>>>backToZeroAfterJob : [ '+backToZeroAfterJob+' ]');
            }
            elapsedJobTime = Math.round(100 * tracePreview(myPreviewLayer,myLayerPath,backToZeroAfterJob))/100;
            elapsedTotalTime += elapsedJobTime;
            timeText += (myLayerName != 'Mixed Jobs LASER_PREVIEW')? myLayerPath[0][0][0]+' >>> Estimated time : ' : MixedJobsNames[i] +' >>> Estimated time : ';
            seconds = Math.floor(60 * (elapsedJobTime - Math.floor(elapsedJobTime)));
            timeText += Math.floor(elapsedJobTime) + ' minutes '+seconds+' seconds';
            writeEstimatedTime(myLayerName,timeText);
        } else {
            timeText  ='No preview for [ '+myLayerPath[0][0][0]+' ]-> no estimated time'
            writeEstimatedTime(myLayerName,timeText);
        }
    }
    seconds = Math.floor(60 * (elapsedTotalTime - Math.floor(elapsedTotalTime)));
    timeText  ='Previewed jobs total estimated time : '+Math.floor(elapsedTotalTime)+ ' minutes '+seconds+' seconds';
    writeEstimatedTime(myLayerName,timeText);
}

function writeEstimatedTime(myLayerName,timeText){
    var myLayer = docObj.layers.getByName(myLayerName);
    var h = Math.round(100 * docObj.height)/100;
    var txtY;
    
    var myTxtItem = myLayer.textFrames.add();
    var texteRange = myTxtItem.textRange;
    myTxtItem.contents = timeText;
    txtY = -h +myLayer.textFrames.length * 50;
    myTxtItem.position = [200, txtY]; // [X, Y]
    texteRange.size = 32;
    //texteRange.font = "Calibri";
    //texteRange.fontStyle = "ExtraBold";
}

function tracePreview(myPreviewLayer,myLayerPath,backToZeroAfterJob) {
    var myStrokeModel;
    var mySpeed;
    var myPathInGCODE;
    var myPathData;
    var elapsedTime = 0;
    var allCoordinates;
    var backToZero = new Array(new Array(0,0));
    
    for (var j=0;j<myLayerPath[0][0][8];j++){
        for (var i=0;i<myLayerPath.length;i++){
            myPathData = myLayerPath[i];
            mySpeed = myPathData[0][7];
            myStrokeModel = myLayerPath[i][0][3];
            myPathInGCODE = myPathData[4];
            allCoordinates = cleanGCODEtoCoordinatesOnly(myPathInGCODE);
            
            elapsedTime += makeTravelPreview(myPreviewLayer,allCoordinates);
            elapsedTime += makeJobPreview(myPreviewLayer,allCoordinates,myStrokeModel,mySpeed);
        }
    }
    if (backToZeroAfterJob){
        elapsedTime += makeTravelPreview(myPreviewLayer,backToZero);
        LASTPREVIEWPOINT.x = 0;
        LASTPREVIEWPOINT.y = 0;
    }

    return elapsedTime;
}

function makeTravelPreview(myPreviewLayer,allCoordinates){
    var myTime = 0;
    var myPathStroke;
    var startPointInMM = new Point();
    var endPointInMM = new Point();
    var startPointInPT = new Array();
    var endPointInPT = new Array();
    
    startPointInMM.x = LASTPREVIEWPOINT.x;
    startPointInMM.y = LASTPREVIEWPOINT.y;
    endPointInMM.x = allCoordinates[0][0];
    endPointInMM.y = allCoordinates[0][1];
    
    var myDistance = euclideanDistance(startPointInMM,endPointInMM);
    //$.writeln('>>>>>myDistance : [ '+myDistance+' ]'); 
    
    if (myDistance > 0){
        //myTime += myDistance / TRAVELSPEED;
        myTime += travelTime(myDistance, TRAVELSPEED, ACCELERATION)
        myPathStroke = myPreviewLayer.pathItems.add();
        myPathStroke.stroked = true;
        myPathStroke.filled = false;
        myPathStroke.strokeColor = TRAVELSTROKE.strokeColor;
        myPathStroke.strokeWidth = TRAVELSTROKE.strokeWidth;
        myPathStroke.strokeDashes = TRAVELSTROKE.strokeDashes;
        myPathStroke.name = myPreviewLayer.pathItems.length;
        
        startPointInPT[0] = startPointInMM.x / mmPointRatio;
        startPointInPT[1] = (startPointInMM.y - TRIMORIGINY) / mmPointRatio;
        
        endPointInPT[0] = endPointInMM.x / mmPointRatio;
        endPointInPT[1] = (endPointInMM.y - TRIMORIGINY) / mmPointRatio;
        
        var startpt = myPathStroke.pathPoints.add();
        startpt.anchor = startPointInPT;
        startpt.leftDirection = startPointInPT;
        startpt.rightDirection = startPointInPT;
        startpt.pointType = PointType.CORNER;
        var endpt = myPathStroke.pathPoints.add();
        endpt.anchor = endPointInPT;
        endpt.leftDirection = endPointInPT;
        endpt.rightDirection = endPointInPT;
        endpt.pointType = PointType.CORNER;
    }
    
    return myTime;
}

function makeJobPreview(myPreviewLayer,allCoordinates,myStrokeModel,mySpeed){
    var myTime = 0;
    var myPathStroke;
    var startPointInMM = new Point();
    var endPointInMM = new Point();
    var startPointInPT = new Array();
    var endPointInPT = new Array();
    
    endPointInMM.x = LASTPREVIEWPOINT.x;
    endPointInMM.y = LASTPREVIEWPOINT.y;
    myPathStroke = myPreviewLayer.pathItems.add();
    myPathStroke.stroked = true;
    myPathStroke.filled = false;
    myPathStroke.strokeColor = myStrokeModel.strokeColor;
    myPathStroke.strokeWidth = myStrokeModel.strokeWidth;
    myPathStroke.strokeDashes = myStrokeModel.strokeDashes;
    myPathStroke.name = myPreviewLayer.pathItems.length;
    
    for (var i=1;i<allCoordinates.length;i++){
        startPointInMM.x = allCoordinates[i-1][0];
        startPointInMM.y = allCoordinates[i-1][1];
        endPointInMM.x = allCoordinates[i][0];
        endPointInMM.y = allCoordinates[i][1];
        
        myDistance = euclideanDistance(startPointInMM,endPointInMM);
        //myTime += myDistance / mySpeed;
        myTime += travelTime(myDistance, mySpeed, ACCELERATION)
        
        startPointInPT[0] = startPointInMM.x / mmPointRatio;
        startPointInPT[1] = (startPointInMM.y - TRIMORIGINY) / mmPointRatio;
        
        endPointInPT[0] = endPointInMM.x / mmPointRatio;
        endPointInPT[1] = (endPointInMM.y - TRIMORIGINY) / mmPointRatio;
        
        var startpt = myPathStroke.pathPoints.add();
        startpt.anchor = startPointInPT;
        startpt.leftDirection = startPointInPT;
        startpt.rightDirection = startPointInPT;
        startpt.pointType = PointType.CORNER;
        var endpt = myPathStroke.pathPoints.add();
        endpt.anchor = endPointInPT;
        endpt.leftDirection = endPointInPT;
        endpt.rightDirection = endPointInPT;
        endpt.pointType = PointType.CORNER;
    }
    LASTPREVIEWPOINT.x = endPointInMM.x;
    LASTPREVIEWPOINT.y = endPointInMM.y;
    
    return myTime;
}

function travelTime(distance_mm, speed_mm_min, accel_mm_s2)
{
    var time;
    var v = speed_mm_min / 60.0; // mm/s
    var a = accel_mm_s2;

    var d_min = (v * v) / a;

    if (distance_mm >= d_min)
    {
        time = (distance_mm / v) + (v / a);
        return time/60;
    }
    else
    {
        var v_peak = Math.sqrt(a * distance_mm);
        time = (2 * v_peak) / a;
        return time/60;
    }
}

function cleanGCODEtoCoordinatesOnly(myPathInGCODE){  
    var allCoordinates = new Array();
    var myPoint = new Array();
    myPoint[0] = LASTPREVIEWPOINT.x;
    myPoint[1] = LASTPREVIEWPOINT.y;
    var myCommandArray = myPathInGCODE.split('\n');
    var myCommand;
    var myCoordinates;
    var yIndex;
    
    for (var i=0;i<myCommandArray.length;i++){
        myCommand = myCommandArray[i];
        if (myCommand.charAt(0) == 'X' || myCommand.charAt(0) == 'Y'){
            myCoordinates = myCommand.split(' ');
            myPoint[0] = parseFloat(myCoordinates[0].substring(1,myCoordinates[0].length));
            myPoint[1] = parseFloat(myCoordinates[1].substring(1,myCoordinates[1].length));
            allCoordinates.push(new Array(myPoint[0],myPoint[1]));
        }
    }
    
    return allCoordinates;    
}

////////////////////////////////////////////////////////// CONVERT PATH TO DASHED ////////////////////////////////////
function dashedPath(myPath, plain, gap){

    var pts = myPath.pathPoints;
    var closed = myPath.closed;
    var poly = buildAdaptivePolyline(myPath);

    if(poly.length < 2){
        return [];
    }

    var distances = buildDistances(poly);
    var totalLength = distances[distances.length-1];

    if(totalLength < (2*plain + gap)){
        return [buildPathPointArray(poly)];
    }

    var dashPeriod = plain + gap;
    var segments = [];
    var pos = 0;

    var maxIter = Math.ceil(totalLength / plain) + 5;
    var iter = 0;

    while(pos < totalLength && iter < maxIter){

        var start = pos;
        var end = pos + plain;

        if(end > totalLength){
            end = totalLength;
        }

        segments.push(extractSection(poly, distances, start, end));
        pos += dashPeriod;
        iter++;
    }

    return segments;
}

function buildAdaptivePolyline(myPath){
    var pts = myPath.pathPoints;
    var n = pts.length;
    var closed = myPath.closed;
    var poly = [];

    if(n < 2){
        return poly;
    }

    var segmentCount = closed ? n : n-1;

    for(var i=0;i<segmentCount;i++){
        var p1 = pts[i];
        var p2 = pts[(i+1) % n];

        var A = p1.anchor;
        var B = p1.rightDirection;
        var C = p2.leftDirection;
        var D = p2.anchor;

        // ajouter point de départ du segment
        if(poly.length == 0){
            poly.push([A[0],A[1]]);
        }

        if(isLine(A,B,C,D)){
            poly.push([D[0],D[1]]);
        }else{
            var steps = computeAdaptiveSteps(A,B,C,D);

            for(var s=1;s<=steps;s++){
                var t = s/steps;
                var P = bezier(A,B,C,D,t);

                poly.push(P);
            }
        }
    }

    return poly;
}

function isLine ( A , B , C , D ){

    return (
    A [ 0 ] == B [ 0 ] &&
    A [ 1 ] == B [ 1 ] &&
    C [ 0 ] == D [ 0 ] &&
    C [ 1 ] == D [ 1 ]
    );
}

function computeAdaptiveSteps(A,B,C,D){
    var l = dist(A,B) + dist(B,C) + dist(C,D);

    var chord = dist(A,D);
    var curvature = l - chord;

    var steps = Math.ceil(BEZIERSTEPS * (1 + curvature/20));

    if(steps < 2) steps = 2;
    if(steps > 50) steps = 50;

    return steps;
}

function dist(A,B){
    var dx = B[0]-A[0];
    var dy = B[1]-A[1];

    return Math.sqrt(dx*dx + dy*dy);
}

function buildDistances(poly){
    var d = [0];

    for(var i=1;i<poly.length;i++){
        var L = dist(poly[i-1],poly[i]);

        d.push(d[i-1] + L);
    }

    return d;
}

function extractSection(poly,distances,dStart,dEnd){
    var pts = [];
    var pStart = pointAtDistance(poly,distances,dStart);
    var pEnd = pointAtDistance(poly,distances,dEnd);

    pts.push(pStart);

    for(var i=1;i<distances.length;i++){
        if(distances[i] > dStart && distances[i] < dEnd){
            pts.push(poly[i]);
        }
    }

    pts.push(pEnd);

    return buildPathPointArray(pts);
}

function pointAtDistance(poly,distances,d){

    for(var i=1;i<distances.length;i++){
        if(distances[i] >= d){
            var A = poly[i-1];
            var B = poly[i];

            var da = distances[i-1];
            var db = distances[i];

            var t = (d-da)/(db-da);

            return [
                A[0] + t*(B[0]-A[0]),
                A[1] + t*(B[1]-A[1])
            ];
        }
    }

    return poly[poly.length-1];
}

function buildPathPointArray(poly){
    var arr = [];

    for(var i=0;i<poly.length;i++){
        var p = poly[i];

        arr.push({
            anchor:[p[0],p[1]],
            leftDirection:[p[0],p[1]],
            rightDirection:[p[0],p[1]]
        });
    }

    return arr;
}