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
var port = process.env.PORT || config.PORT;
var app = express();

var input = __dirname + '/temp/1.pdf';

pdf2img.setOptions({
    type: 'jpg', // png or jpg, default jpg 
    size: 1024, // default 1024 
    density: 600 // default 600                                // convert selected page, default null (if null given, then it will convert all pages) 
});

pdf2img.convert(input, function(err, info) {
    if (err) console.log(err)
    else console.log(info);
});
return;
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
        unoconv.convert(file.fullPath, 'pdf', function(err, result) {
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

app.post('/convert_from_pdf', jsonParser, function(req, res) {
    var data = req.body;
    var convertedFiles = [];
    data.forEach(function(element, index, array) {
        var file = {

        };
        file.fullPath = config.TEMP + element;
        file.convertTo = element.split('.')[0].trim() + ".jpg";
        file.writeTo = config.TEMP + file.convertTo;
        var pdf = scissors(file.fullPath);
        pdf.getNumPages().then(function(pages) {
                process(1, pages);
            },
            function(fail) {
                console.log(fail);
            });

        function process(i, pages) {
            if (pages < i) {
                convertToJpg(1, pages);
                return;
            }
            var pdf = scissors(file.fullPath);

            var fullPath = config.TEMP + 'pdfs/' + i + '.pdf';
            pdf.pages(i).pdfStream().pipe(fs.createWriteStream(fullPath))
                .on('finish', function() {
                    i = i + 1;
                    console.log(i);
                    process(i, pages);
                }).on('error', function(err) {
                    console.log(err);
                });
        }
    });

    function convertToJpg(i, total) {
        if (i > total) {
            archive(total, function(result) {
                res.download(result);
            });
            return;
        }
        var fullPath = config.TEMP + 'pdfs/' + i + '.pdf';
        var writeTo = config.TEMP + 'pdfs/' + i + '.jpg';
        unoconv.convert(fullPath, 'jpg', { port: 2002 }, function(err, result) {
            if (err) {
                console.log(err);
                return;
            }
            fs.writeFileSync(writeTo, result);
            i = i + 1;
            convertToJpg(i, total);
        });
    }

    function archive(total, callback) {
        console.log('archive');
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
        for (i = 1; i <= total; i++) {
            archive.append(fs.createReadStream(config.TEMP + 'pdfs/' + i + '.jpg'), { name: i + '.jpg' })
        }
        archive.finalize();
    }

    function convert(fullPath, writeTo, convertTo, callback) {
        unoconv.convert(fullPath, 'jpg', function(err, result) {
            // result is returned as a Buffer
            fs.writeFileSync(writeTo, result);


        });
    }

});


app.listen(port);