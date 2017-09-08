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
var port = process.env.PORT || config.PORT;
var app = express();

app.use(compression());
app.use(minify());
app.use('/', express.static(__dirname + '/public'));
app.set('view engine', 'ejs');

app.get('/', function(req, res) {
    res.render('index');
});

app.get('/word_to_pdf', function(req, res) {
    res.render('word_to_pdf');
});

app.post('/word_to_pdf', multipartyMiddleware, function(req, res) {
    fs.readFile(req.files.file.path, function(err, data) {
        var file = req.files.file;
        file.fullPath = config.TEMP + file.name;
        if (file.type !== 'application/msword' && file.type !== 'application/vnd.openxmlformats-officedocument.wordprocessingml.document') {
            res.sendStatus(404);
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
            fs.writeFile(file.writeTo, result);
            convertedFiles.push(file.writeTo);
        });
    });
    var output = fs.createWriteStream(__dirname + '/output.zip');
    var archive = archiver('zip', {
        gzip: true,
        zlib: { level: 9 } // Sets the compression level.
    });

    archive.on('error', function(err) {
        console.log(err);
    });
    output.on('close', function() {
        res.sendFile(__dirname + '/output.zip');
    });
    archive.pipe(output);
    convertedFiles.forEach(function(file, index, array) {
        archive.append(fs.createReadStream(file), { name: file.split(config.TEMP)[1].trim() });
    });
    archive.finalize();
});


app.listen(port);
