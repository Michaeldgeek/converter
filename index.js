var express = require("express");
var config = require("./config");
var compression = require("compression");
var minify = require("express-minify");
var multiparty = require("connect-multiparty");
var fs = require("fs");
var bodyParser = require("body-parser");
const zlib = require("zlib");
var multipartyMiddleware = multiparty();
var archiver = require("archiver");
var pdf2img = require("pdf2img");
var slugify = require("slugify");
var shell = require("shelljs");
var port = process.env.PORT || config.PORT;
var app = express();

app.use(compression());
app.use(minify());
app.use("/", express.static(__dirname + "/public"));
app.use("/downloads/", express.static(__dirname + "/downloads"));
app.set("view engine", "ejs");

app.get("/", function (req, res) {
    res.render("index");
});

app.get("/word_to_pdf", function (req, res) {
    res.render("word_to_pdf");
});

app.get("/image_to_pdf", function (req, res) {
    res.render("image_to_pdf");
});

app.get("/powerpoint_to_pdf", function (req, res) {
    res.render("powerpoint_to_pdf");
});

app.get("/excel_to_pdf", function (req, res) {
    res.render("excel_to_pdf");
});

app.get("/pdf_to_jpg", function (req, res) {
    res.render("pdf_to_jpg");
});

app.get("/pdf_to_word", function (req, res) {
    res.render("pdf_to_word");
});

app.post("/download", function (req, res) {
    var file = req.get("file");
    var response = fs.existsSync(__dirname + "/" + config.DOWNLOADS + file);
    if (response) {
        res.send({
            path: config.DOMAIN + "/" + config.DOWNLOADS + file
        });
    } else {
        res.status(404);
        res.send("error");
    }
});

app.post("/word_to_pdf", multipartyMiddleware, function (req, res) {
    fs.readFile(req.files.file.path, function (err, data) {
        var file = req.files.file;
        file.fullPath =
            config.TEMP +
            slugify(file.name, {
                remove: /[*+~()'"!:@]/g
            });
        file.type = file.type.toLowerCase();
        if (
            file.type !== "application/msword" &&
            file.type !==
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        ) {
            res.status(404);
            res.send("Select a docx file");
            return;
        }
        fs.writeFile(file.fullPath, data, function (err) {
            if (err) {
                console.log(err);
                res.send("error");
                return;
            }
            res.send("ok");
        });
    });
});

app.post("/powerpoint_to_pdf", multipartyMiddleware, function (req, res) {
    fs.readFile(req.files.file.path, function (err, data) {
        var file = req.files.file;
        file.fullPath =
            config.TEMP +
            slugify(file.name, {
                remove: /[*+~()'"!:@]/g
            });
        file.type = file.type.toLowerCase();
        if (
            file.type !== "application/vnd.ms-powerpoint" &&
            file.type !==
            "application/vnd.openxmlformats-officedocument.presentationml.presentation" &&
            file.type !== ".ppt" &&
            file.type !== ".pptx"
        ) {
            res.status(404);
            res.send("Select a powerpoint file");
            return;
        }
        fs.writeFile(file.fullPath, data, function (err) {
            if (err) {
                console.log(err);
                res.send("error");
                return;
            }
            res.send("ok");
        });
    });
});

app.post("/excel_to_pdf", multipartyMiddleware, function (req, res) {
    fs.readFile(req.files.file.path, function (err, data) {
        var file = req.files.file;
        file.fullPath =
            config.TEMP +
            slugify(file.name, {
                remove: /[*+~()'"!:@]/g
            });
        file.type = file.type.toLowerCase();
        if (
            file.type !==
            "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" &&
            file.type !== "application/vnd.ms-excel" &&
            file.type !== ".xlsx" &&
            file.type !== ".xls"
        ) {
            res.status(404);
            res.send("Select an Excel document");
            return;
        }
        fs.writeFile(file.fullPath, data, function (err) {
            if (err) {
                console.log(err);
                res.send("error");
                return;
            }
            res.send("ok");
        });
    });
});

app.post("/image_to_pdf", multipartyMiddleware, function (req, res) {
    fs.readFile(req.files.file.path, function (err, data) {
        var file = req.files.file;
        file.fullPath =
            config.TEMP +
            slugify(file.name, {
                remove: /[*+~()'"!:@]/g
            });
        file.type = file.type.toLowerCase();
        if (
            file.type !== "image/jpg" &&
            file.type !== "image/jpeg" &&
            file.type !== "image/png"
        ) {
            res.status(404);
            res.send("Select an image  file");
            return;
        }
        fs.writeFile(file.fullPath, data, function (err) {
            if (err) {
                res.status(404);
                res.send("error");
                return;
            }
            res.send("ok");
        });
    });
});

app.post("/pdf_to_jpg", multipartyMiddleware, function (req, res) {
    fs.readFile(req.files.file.path, function (err, data) {
        var file = req.files.file;
        file.fullPath =
            config.TEMP +
            slugify(file.name, {
                remove: /[*+~()'"!:@]/g
            });
        file.type = file.type.toLowerCase();
        if (file.type !== "application/pdf") {
            res.status(404);
            res.send("Select a pdf document");
            return;
        }
        fs.writeFile(file.fullPath, data, function (err) {
            if (err) {
                console.log(err);
                res.send("error");
                return;
            }
            res.send("ok");
        });
    });
});

app.post("/pdf_to_word", multipartyMiddleware, function (req, res) {
    fs.readFile(req.files.file.path, function (err, data) {
        var file = req.files.file;
        file.fullPath =
            config.TEMP +
            slugify(file.name, {
                remove: /[*+~()'"!:@]/g
            });
        file.type = file.type.toLowerCase();
        if (file.type !== "application/pdf") {
            res.status(404);
            res.send("Select a pdf document");
            return;
        }
        fs.writeFile(file.fullPath, data, function (err) {
            if (err) {
                console.log(err);
                res.send("error");
                return;
            }
            res.send("ok");
        });
    });
});

var jsonParser = bodyParser.json();
app.post("/convert_to_pdf", jsonParser, function (req, res) {
    var data = req.body;
    var header = req.header("save-as");
    data.forEach(function (element) {
        element = slugify(element, {
            remove: /[*+~()'"!:@]/g
        });
        var file = {};
        file.fullPath = config.TEMP + element;
        file.convertTo = element.split(".")[0].trim() + ".pdf";
        var code = shell.exec(
            "sudo libreoffice  --convert-to pdf --outdir ./downloads " +
            file.fullPath +
            ""
        ).code;
        if (code !== 0) {
            res.status(404);
            res.send("error");
            return;
        }
        var output = fs.createWriteStream(
            __dirname + "/" + config.DOWNLOADS + header
        );
        var archive = archiver("zip", {
            gzip: true,
            zlib: {
                level: 9
            } // Sets the compression level.
        });
        archive.on("error", function () {
            res.status(404);
            res.send("error");
            return;
        });

        output.on("close", function () {
            res.send({
                path: config.DOMAIN + "/" + config.DOWNLOADS + header
            });

        });
        archive.pipe(output);
        archive.append(
            fs.createReadStream(__dirname + "/" + config.DOWNLOADS + file.convertTo), {
                name: file.convertTo
            }
        );
        archive.finalize();
    });
});



app.post("/convert_pdf_to_word", jsonParser, function (req, res) {
    var data = req.body;
    var header = req.header("save-as");
    data.forEach(function (element) {
        element = slugify(element, {
            remove: /[*+~()'"!:@]/g
        });
        var file = {};
        file.fullPath = config.TEMP + element;
        file.convertTo = element.split(".")[0].trim() + ".doc";
        var code = shell.exec(
            "sudo libreoffice --convert-to doc  --infilter='writer_pdf_import' --outdir ./downloads " +
            file.fullPath +
            ""
        ).code;
        if (code !== 0) {
            res.status(404);
            res.send("error");
            return;
        }
        var output = fs.createWriteStream(
            __dirname + "/" + config.DOWNLOADS + header
        );
        var archive = archiver("zip", {
            gzip: true,
            zlib: {
                level: 9
            } // Sets the compression level.
        });
        archive.on("error", function () {
            res.status(404);
            res.send("error");
            return;
        });

        output.on("close", function () {
            res.send({
                path: config.DOMAIN + "/" + config.DOWNLOADS + header
            });
        });
        archive.pipe(output);
        archive.append(
            fs.createReadStream(__dirname + "/" + config.DOWNLOADS + file.convertTo), {
                name: file.convertTo
            }
        );
        archive.finalize();
    });
});

app.post("/convert_pdf_to_jpg", jsonParser, function (req, res) {
    var data = req.body;
    var header = req.header("save-as");
    data.forEach(function (element) {
        element = slugify(element, {
            remove: /[*+~()'"!:@]/g
        });
        var file = {};
        file.fullPath = config.TEMP + element;
        file.convertTo = element.split(".")[0].trim() + ".jpg";
        pdf2img.setOptions({
            type: "jpg", // png or jpg, default jpg
            size: 1024, // default 1024
            density: 600, // default 600
            outputdir: __dirname + "/" + config.DOWNLOADS + "images", // output folder, default null (if null given, then it will create folder name same as file name)
            outputname: element.split(".")[0].trim(),
            page: null // convert selected page, default null (if null given, then it will convert all pages)
        });
        pdf2img.convert(file.fullPath, function (err, info) {
            if (err) {
                res.status(404);
                res.send("An error occured");
                return;
            } else {
                var output = fs.createWriteStream(
                    __dirname + "/" + config.DOWNLOADS + header
                );
                var archive = archiver("zip", {
                    gzip: true,
                    zlib: {
                        level: 9
                    } // Sets the compression level.
                });
                archive.on("error", function (e) {
                    console.log(e);
                    res.status(404);
                    res.send("error");
                    return;
                });

                output.on("close", function () {
                    res.send({
                        path: config.DOMAIN + "/" + config.DOWNLOADS + header
                    });
                });
                archive.pipe(output);
                info.message.forEach(element => {
                    archive.append(
                        fs.createReadStream(element.path), {
                            name: element.name
                        }
                    );
                });

                archive.finalize();
            }
        });
    });
});

app.listen(port);