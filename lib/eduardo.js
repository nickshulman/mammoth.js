var index = require("./index");
var fs = require('fs');
var path = require('path');

var filesToConvert = [];

fs.readdir("C:\\Users\\Eduardo\\repos\\pwiz\\pwiz_tools\\Skyline\\Documentation\\Tutorials", function(err, files) {
    if (err) {
        throw err;
    }
    // eslint-disable-next-line brace-style
    else {
        files = files.filter(
            function(file) {
                return file.endsWith(".docx");
            });
        // eslint-disable-next-line no-console
        filesToConvert.apply(files);
    }
});

filesToConvert.forEach(function(file) {
    var outputFileName = path.parse(file).name + ".html"

    var result = index.convertToHtml({path: file});

    result.then(function(result) {
        // eslint-disable-next-line no-console
        console.log(result);

        fs.writeFile("/Users/eduardoarmendariz/Desktop/Skyline Document Processing/Skyline Absolute Quantification.html", result.value, function(err) {
            if (err) {
                throw err;
            }
        });
    });
});

// var result = index.convertToHtml({path: "/Users/eduardoarmendariz/Desktop/Skyline Document Processing/Skyline Absolute Quantification.docx"});
// result.then(function(result) {
//     // eslint-disable-next-line no-console
//     console.log(result);
//
//     fs.writeFile("/Users/eduardoarmendariz/Desktop/Skyline Document Processing/Skyline Absolute Quantification.html", result.value, function(err) {
//         // In case of a error throw err.
//         if (err) {
//             throw err;
//         }
//     });
// });