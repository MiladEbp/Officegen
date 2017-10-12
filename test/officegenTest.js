"use strict";
exports.__esModule = true;
var fs_1 = require("fs");
var officeGen = require('officegen');
var dir = "outStream/";
describe('officegen', function () {
    it('Add Header in File', function (done) {
        var docxFile = officeGen('docx');
        var header = docxFile.getHeader().createP({ align: 'center' });
        header.addText('Unit Test For officegen', { color: 'black' });
        var fileName = 'header.docx';
        var out = fs_1.createWriteStream(dir + fileName);
        docxFile.generate(out);
        out.on('close', function () {
            console.log('Create file of ' + fileName + ' in directory');
            done();
            return officeGen.Promise = global.Promise;
        });
    });
    it('Add Description in File', function (done) {
        var docxFile = officeGen('docx');
        var createP = docxFile.createP({ align: 'center' });
        createP.addText('این فایل برای تست است .', { color: 'red', font_face: 'B Nazanin' });
        var fileName = 'description.docx';
        var out = fs_1.createWriteStream(dir + fileName);
        docxFile.generate(out);
        out.on('close', function () {
            console.log('Create file of ' + fileName + ' in directory');
            done();
            return officeGen.Promise = global.Promise;
        });
    });
});
//# sourceMappingURL=officegenTest.js.map