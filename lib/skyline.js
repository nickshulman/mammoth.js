var mammoth = require("./index");
var fs = require('fs');
var path = require('path');
var beautify = require("js-beautify")
var _ = require('underscore');

var baseTutorialPath = "C:\\Users\\Eduardo\\repos\\pwiz\\pwiz_tools\\Skyline\\Documentation\\Tutorials";
var baseOutputPath = "C:\\Users\\Eduardo\\repos\\data\\SkylineDocs\\generated-docs";
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


fs.rmSync(baseOutputPath, {recursive: true, force: true});
convertAll()

function convertAll(){
    conversions.forEach(function(conversion){
        convertDirectory(conversion)
    })
}

function convertDirectory(conversion){
    fs.mkdirSync(conversion.outputPath);
    fs.readdir(conversion.sourcePath, function(err, files) {
        if (err) {
            throw err;
        }
        files.filter(
            function(file) {
                return file.endsWith(".docx");
            }).forEach(function(file) {
            convertDocument(conversion, file);
        });

    });

}

function convertDocument(conversion, sourceDocument) {
    var documentName = path.parse(sourceDocument).name;
    var outputFileName = documentName + ".html";
    var outputFolder = conversion.outputPath + "/" + documentName;
    var outputFile = outputFolder + "\\" + outputFileName;
    var imageCounter = 0;
    fs.mkdirSync(outputFolder);

    var options = {
        styleMap: [
            "p[style-name='Title'] => h1.document-title",
            "p[style-name='Bibliography'] => p.bibliography"
        ],
        convertImage: mammoth.images.imgElement(function(image) {
            if (!(image.contentType in imageContentTypeExtensionMap)) {
                throw "Unknown image type";
            }
            var extension = imageContentTypeExtensionMap[image.contentType];
            var imageName = "image-" +imageCounter++ +extension;

            return image.readAsBuffer().then(function(imageBuffer) {
                fs.writeFile(outputFolder + "\\" +imageName, imageBuffer, function(err) {
                    if (err) {
                        throw err;
                    }
                });
                return {
                    src: imageName
                };
            });
        })
    };

    var result = mammoth.convertToHtml({path: conversion.sourcePath + "/" + sourceDocument}, options);
    result.then(function(result) {
        console.log(sourceDocument)
        console.log(result.messages)
        var wrappedHtml = `<html><head><link rel="stylesheet" type="text/css" href="${conversion.cssPath}"></head><body>${result.value}</body></html>`
        var formattedHtml = formatHtml(wrappedHtml);

        fs.writeFile(outputFile, formattedHtml, function(err) {
            if (err) {
                throw err;
            }
        });
    });
}

function formatHtml(html){
    return beautify.html_beautify(html, htmlPrettyPrintOptions);;
}

/*
 *  Issues:
 *  Tables within a list do not respect boundaries - Fixed by implementing table indentation
 *  Bibliography style not applied to all - Implemented bibliography class but some documents dont use the style
 *  Styling of output document - Used beautify-js to format html and add line wrap limits
 */