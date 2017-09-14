// there are some raw file path in your codes
var express = require('express');
var config = require('./config');
var compression = require('compression');
var minify = require('express-minify');
var multiparty = require('connect-multiparty');
var unoconv = require('unoconv2');
var fs = require('fs');
var bodyParser = require('body-parser');
const zlib = require('zlib');
const gzip = zlib.createGzip();
var multipartyMiddleware = multiparty();
var archiver = require('archiver');
var scissors = require('scissors');
var hummus = require('hummus');
var pdf2img = require('pdf2img');
var rmdir = require('rmdir');
var mv = require('mv');
var shell = require('shelljs');
var HtmlDocx = require('html-docx-js');
var port = process.env.PORT || config.PORT;
var app = express();

app.use(compression());
app.use(minify());
app.use('/', express.static(__dirname + '/public'));
app.set('view engine', 'ejs');


unoconv.listen({ port: 2002 });
app.get('/', function(req, res) {
    res.render('index');
});

app.get('/word_to_pdf', function(req, res) {
    res.render('word_to_pdf');
});

app.get('/jpg_to_pdf', function(req, res) {
    res.render('jpg_to_pdf');
});

app.get('/powerpoint_to_pdf', function(req, res) {
    res.render('powerpoint_to_pdf');
});

app.get('/excel_to_pdf', function(req, res) {
    res.render('excel_to_pdf');
});

app.get('/pdf_to_jpg', function(req, res) {
    res.render('pdf_to_jpg');
});

app.get('/pdf_to_word', function(req, res) {
    res.render('pdf_to_word');
});


app.post('/download', function(req, res) {
    var response = fs.existsSync('./output.zip');
    if (response) {
        res.download('./output.zip');
    } else {
        res.status(404);
        res.send('error');
    }
});

app.post('/word_to_pdf', multipartyMiddleware, function(req, res) {
    fs.readFile(req.files.file.path, function(err, data) {
        var file = req.files.file;
        file.fullPath = config.TEMP + file.name;
        if (file.type !== 'application/msword' && file.type !== 'application/vnd.openxmlformats-officedocument.wordprocessingml.document') {
            res.status(404);
            res.send("Select a docx file");
            return;
        }
        fs.writeFile(file.fullPath, data, function(err) {
            if (err) {
                console.log(err);
                res.send('error');
                return;
            }
            res.send('ok');
        });

    });
});

app.post('/powerpoint_to_pdf', multipartyMiddleware, function(req, res) {
    fs.readFile(req.files.file.path, function(err, data) {
        var file = req.files.file;
        file.fullPath = config.TEMP + file.name;
        if ((file.type !== 'application/vnd.ms-powerpoint' && file.type !== 'application/vnd.openxmlformats-officedocument.presentationml.presentation') && (file.type !== '.ppt' && file.type !== '.pptx')) {
            res.status(404);
            res.send("Select a powerpoint file");
            return;
        }
        fs.writeFile(file.fullPath, data, function(err) {
            if (err) {
                console.log(err);
                res.send('error');
                return;
            }
            res.send('ok');
        });

    });
});

app.post('/excel_to_pdf', multipartyMiddleware, function(req, res) {
    fs.readFile(req.files.file.path, function(err, data) {
        var file = req.files.file;
        file.fullPath = config.TEMP + file.name;
        if ((file.type !== 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' && file.type !== 'application/vnd.ms-excel') && (file.type !== '.xlsx' && file.type !== '.xls')) {
            res.status(404);
            res.send("Select an Excel document");
            return;
        }
        fs.writeFile(file.fullPath, data, function(err) {
            if (err) {
                console.log(err);
                res.send('error');
                return;
            }
            res.send('ok');
        });

    });
});

app.post('/jpg_to_pdf', multipartyMiddleware, function(req, res) {
    fs.readFile(req.files.file.path, function(err, data) {
        var file = req.files.file;
        file.fullPath = config.TEMP + file.name;
        file.type = file.type.toLowerCase();
        if (file.type !== 'image/jpg' && (file.type !== 'image/jpeg' && file.type !== 'image/png')) {
            res.status(404);
            res.send("Select an image  file");
            return;
        }
        fs.writeFile(file.fullPath, data, function(err) {
            if (err) {
                console.log(err);
                res.send('error');
                return;
            }
            res.send('ok');
        });

    });
});

app.post('/pdf_to_jpg', multipartyMiddleware, function(req, res) {
    fs.readFile(req.files.file.path, function(err, data) {
        var file = req.files.file;
        file.fullPath = config.TEMP + file.name;
        if (file.type !== 'application/pdf') {
            res.status(404);
            res.send("Select a pdf document");
            return;
        }
        fs.writeFile(file.fullPath, data, function(err) {
            if (err) {
                console.log(err);
                res.send('error');
                return;
            }
            res.send('ok');
        });

    });
});

app.post('/pdf_to_word', multipartyMiddleware, function(req, res) {
    fs.readFile(req.files.file.path, function(err, data) {
        var file = req.files.file;
        file.fullPath = config.TEMP + file.name;
        if (file.type !== 'application/pdf') {
            res.status(404);
            res.send("Select a pdf document");
            return;
        }
        fs.writeFile(file.fullPath, data, function(err) {
            if (err) {
                console.log(err);
                res.send('error');
                return;
            }
            res.send('ok');
        });

    });
});

var jsonParser = bodyParser.json()
app.post('/convert_to_pdf', jsonParser, function(req, res) {
    var data = req.body;
    var convertedFiles = [];
    data.forEach(function(element, index, array) {
        var file = {

        };
        file.fullPath = config.TEMP + element;
        file.convertTo = element.split('.')[0].trim() + ".pdf";
        file.writeTo = config.TEMP + file.convertTo;
        unoconv.convert(file.fullPath, 'pdf', { port: 2002 }, function(err, result) {
            // result is returned as a Buffer
            fs.writeFileSync(file.writeTo, result);
            var output = fs.createWriteStream(__dirname + '/output.zip');
            var archive = archiver('zip', {
                gzip: true,
                zlib: { level: 9 } // Sets the compression level.
            });
            archive.on('error', function(err) {
                console.log(err);
            });
            output.on('close', function() {
                res.download(__dirname + '/output.zip');
            });
            archive.pipe(output);
            archive.append(fs.createReadStream(file.writeTo), { name: file.convertTo });
            archive.finalize();
        });
    });

});

app.post('/convert_from_pdf_word', jsonParser, function(req, res) {
    var data = req.body;
    var convertedFiles = [];
    data.forEach(function(element, index, array) {
        var file = {

        };
        file.fullPath = config.TEMP + element;
        file.convertTo = element.split('.')[0].trim() + ".docx";
        file.writeTo = config.TEMP + file.convertTo;
        fs.rename('/var/www/converter/temp/' + element, '/var/www/converter/temp/' + element.replace(/\s/g, ''), function(err) {
            if (err) {
                res.status(404);
                res.send("Error occured");
                console.log(err);
                return;
            }
            element = element.replace(/\s/g, '');
            file.fullPath = config.TEMP + element;
            file.convertTo = element.split('.')[0].trim() + ".docx";
            file.writeTo = config.TEMP + file.convertTo;
            var code = shell.exec('sudo mv /var/www/converter/temp/' + element + ' ' + config.LIBRE_OFFICE_PATH + '').code;
            if (code !== 0) {
                res.status(404);
                res.send("Error occured");
                console.log(err);
                return;
            }
            var code = shell.exec("sudo /snap/bin/libreoffice --infilter='writer_pdf_import' --convert-to doc '" + config.LIBRE_OFFICE_PATH + element + "'").code;
            return;
            var archive = archiver('zip', {
                gzip: true,
                zlib: { level: 9 } // Sets the compression level.
            });
            archive.on('error', function(err) {
                console.log(err);
            });
            output.on('close', function() {
                res.download(__dirname + '/output.zip');
            });
            archive.pipe(output);
            archive.append(fs.createReadStream(file.writeTo), { name: file.convertTo });
            archive.finalize();
        });

    });

});

app.post('/convert_from_pdf', jsonParser, function(req, res) {
    var data = req.body;
    var convertedFiles = [];
    data.forEach(function(element, index, array) {
        var file = {

        };
        file.fullPath = config.TEMP + element;
        file.convertTo = element.split('.')[0].trim() + ".jpg";
        file.writeTo = config.TEMP + file.convertTo;
        pdf2img.setOptions({
            type: 'jpg', // png or jpg, default jpg 
            size: 1024, // default 1024 
            density: 600, // default 600 
            outputdir: config.TEMP + 'pdfs', // output folder, default null (if null given, then it will create folder name same as file name) 
            page: null // convert selected page, default null (if null given, then it will convert all pages) 
        });

        pdf2img.convert(file.fullPath, function(err, info) {
            if (err) {
                console.log(err);
                res.status(404);
                res.send('An error occured');
                return;
            } else {
                archive(config.TEMP + "pdfs/", function(result) {
                    res.download(result);
                    rmdir(config.TEMP + "pdfs/", function(err, dirs, files) {});
                });
            };
        });
    });

    function archive(directory, callback) {
        var output = fs.createWriteStream(__dirname + '/output.zip');
        var archive = archiver('zip', {
            gzip: true,
            zlib: { level: 9 } // Sets the compression level.
        });
        archive.on('error', function(err) {
            console.log(err);
        });
        output.on('close', function() {
            callback(__dirname + '/output.zip');
        });
        archive.pipe(output);
        archive.directory(directory, false);
        archive.finalize();
    }
});


app.listen(port);