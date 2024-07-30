var mammoth = require("./index");
var fs = require('fs');
var path = require('path');
var beautify = require("js-beautify")
var _ = require('underscore');
var exec = require('child_process');

//Requires installation of emf2svg tool. see: https://github.com/metanorma/libemf2svg
var emf2svgPath = "C:\\Users\\Eduardo\\repos\\libemf2svg_windows\\Debug\\emf2svg-conv.exe";
var baseTutorialPath = "C:\\Users\\Eduardo\\repos\\pwiz_unchanged\\pwiz_tools\\Skyline\\Documentation\\Tutorials";
var baseOutputPath = "C:\\Users\\Eduardo\\repos\\pwiz\\pwiz_tools\\Skyline\\Documentation\\Tutorials";
var baseCssRelativePath = "..\\..\\SkylineStyles.css";
var htmlPrettyPrintOptions = options = { wrap_line_length: 150};
var conversions = [
    {
        sourcePath: baseTutorialPath,
        outputPath: baseOutputPath,
        cssPath: baseCssRelativePath
    },
    {
        sourcePath: baseTutorialPath + "\\Chinese",
        outputPath: baseOutputPath + "\\Chinese",
        cssPath: "..\\" + baseCssRelativePath
    },
    {
        sourcePath: baseTutorialPath + "\\Chinese\\outgoing",
        outputPath: baseOutputPath + "\\Chinese\\outgoing",
        cssPath: "..\\..\\"  + baseCssRelativePath
    },
    {
        sourcePath: baseTutorialPath + "\\Japanese",
        outputPath: baseOutputPath + "\\Japanese",
        cssPath: "..\\" + baseCssRelativePath
    },
    {
        sourcePath: baseTutorialPath + "\\Japanese\\outgoing",
        outputPath: baseOutputPath + "\\Japanese\\outgoing",
        cssPath: "..\\..\\"  + baseCssRelativePath
    },
]

var imageContentTypeExtensionMap = {
    'image/png': '.png',
    'image/x-emf': '.emf',
    'image/x-wmf': '.wmf',
    'image/jpeg': '.jpeg'
}

var ignoredWarnings = [
    "Unrecognised paragraph style: 'List Paragraph' (Style ID: ListParagraph)", //verified no additional formatting needed
    "Image of type image/x-emf is unlikely to display in web browsers", //post processor conversion
    "Image of type image/x-wmf is unlikely to display in web browsers", //post processor conversion
    "Unrecognised run style: 'Internet Link' (Style ID: InternetLink)", //verified no additional formatting needed
    "Unrecognised paragraph style: 'No Spacing' (Style ID: NoSpacing)", //verified no additional formatting needed .. maybe verify again?
    "Unrecognised paragraph style: 'Body Text' (Style ID: BodyText)", //verified no additional formatting needed
    "Unrecognised paragraph style: 'Normal (Web)' (Style ID: NormalWeb)", //verified no additional formatting needed
    "Unrecognised paragraph style: 'Default' (Style ID: Default)", //verified no additional formatting needed

    "An unrecognised element was ignored: v:stroke", //documented but no solution possible
    "An unrecognised element was ignored: v:path", //documented but no solution possible
    "An unrecognised element was ignored: v:oval", //documented but no solution possible
    "An unrecognised element was ignored: office-word:anchorlock", //documented but no solution possible
    "An unrecognised element was ignored: {urn:schemas-microsoft-com:office:office}OLEObject", //documented but maybe solution possible
    "An unrecognised element was ignored: {http://schemas.openxmlformats.org/officeDocument/2006/math}oMathPara" //documented but maybe solution possible
]

//stateful
var imagesToConvert = [];
fs.rmSync(baseOutputPath, {recursive: true, force: true});
convertAll()

function convertAll(){
    var documentConversionPromises = [];
    conversions.forEach(function(conversion){
        documentConversionPromises.push(...convertDirectory(conversion));
    })

    //converting images at the time of document generation causes file conflict issues
    //to avoid this we convert all the images once the document generation is complete
    Promise.all(documentConversionPromises)
        .then(function(results){
            imagesToConvert.forEach(function(image){
                convertToSvg(image)
            })
        });
}

function convertDirectory(conversion){
    var documentConversionPromises = [];
    fs.mkdirSync(conversion.outputPath);
    var files = fs.readdirSync(conversion.sourcePath);
    files.filter(
        function(file) {
            return file.endsWith(".docx") && !file.includes("~$");
        }).forEach(function(file) {
            documentConversionPromises.push(convertDocument(conversion, file));
        });
    return documentConversionPromises;
}

function convertDocument(conversion, sourceDocument) {
    var documentName = path.parse(sourceDocument).name;
    var outputFileName = documentName + ".html";
    var outputFolder = conversion.outputPath + "\\" + documentName;
    var outputFile = outputFolder + "\\" + outputFileName;
    var imageCounter = 0;
    fs.mkdirSync(outputFolder);

    var options = {
        styleMap: [
            "p[style-name='Title'] => h1.document-title",
            "p[style-name='Bibliography'] => p.bibliography:fresh",
            "p[style-name='Bibliography1'] => p.bibliography:fresh",
            "p[style-name='Subtitle'] => p.subtitle",
            "r[style-name='Subtle Emphasis'] => strong",
        ],
        convertImage: mammoth.images.imgElement(function(image) {
            if (!(image.contentType in imageContentTypeExtensionMap)) {
                throw "Unknown image type";
            }
            var extension = imageContentTypeExtensionMap[image.contentType];
            var imageName = "image-" +imageCounter++;
            var imageFileName = imageName +extension;
            var imageLinkName = imageFileName;
            //rename link to svg as we are converting these files
            if(extension === ".emf" || extension === ".wmf"){
                imageLinkName = imageName + ".svg"
                imagesToConvert.push(outputFolder + "\\" +imageFileName)
            }
            return image.readAsBuffer().then(function(imageBuffer) {
                fs.writeFileSync(outputFolder + "\\" +imageFileName, imageBuffer);
                return {
                    src: imageLinkName
                };
            })
        })
    };

    return mammoth.convertToHtml({path: conversion.sourcePath + "/" + sourceDocument}, options)
        .then(function(result) {
            var messages = result.messages.filter(function(message) {
                return !ignoredWarnings.includes(message.message)
            })
            if(messages !== undefined && messages.length > 0){
                console.log("Found warnings/errors converting document " +sourceDocument)
                console.log(messages);
            }
            var wrappedHtml = `<html><head><link rel="stylesheet" type="text/css" href="${conversion.cssPath}"></head><body>${result.value}</body></html>`
            var formattedHtml = formatHtml(wrappedHtml);

            fs.writeFileSync(outputFile, formattedHtml)
        });
}

function formatHtml(html){
    return beautify.html_beautify(html, htmlPrettyPrintOptions);;
}

function convertToSvg(sourceFile){
    var sourceFileName = path.parse(sourceFile).name;
    var outputFile = path.parse(sourceFile).dir + "\\" +sourceFileName + ".svg"
    var commandArgs = [
        '--input', sourceFile,
        '--output', outputFile
    ]
    exec.execFile(emf2svgPath, commandArgs, function (err, stdout, stderr) {
            if (err) {
                console.log("Could not convert image " +sourceFile)
                if (stderr) console.log(`stderr: ${stderr}`);
                if (stdout) console.log(`stderr: ${stdout}`);
                console.log(`Run command manually: ${emf2svgPath} --verbose --input '${sourceFile}' --output '${outputFile}'`);
            } else {
                fs.rm(sourceFile, function(err, stderr) {
                    if(err){
                        throw err;
                    }
                })
            }
        });

}
/*
 *  Issues:
 *  Tables within a list do not respect boundaries - Fixed by implementing table indentation
 *  Bibliography style not applied to all - Implemented bibliography class but some documents dont use the style
 *  Styling of output document - Used beautify-js to format html and add line wrap limits
 */