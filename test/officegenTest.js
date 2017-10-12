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
            console.log('OutPut is : \nCreate file of ' + fileName + ' in directory');
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
    it('Add Table in File', function (done) {
        var docxFile = officeGen('docx');
        var table = [[{ val: 'نام', opts: { rtl: true, font_face: 'B Nazanin', color: 'block', shd: { fill: '7F7F7F', themeFill: 'text1', themeFillTint: '10' } } },
                { val: 'نام خانوادگی', opts: { rtl: true, font_face: 'B Nazanin', color: 'block', shd: { fill: '7F7F7F', themeFill: 'text1', themeFillTint: '80' } } }],
            [{ val: 'عرفان', opts: { rtl: true, font_face: 'B Nazanin', color: 'red', shd: { fill: '92CDDC', themeFill: 'text1', themeFillTint: '80' } } },
                { val: 'تاواتاو', opts: { rtl: true, font_face: 'B Nazanin', color: 'red', shd: { fill: '92CDDC', themeFill: 'text1', themeFillTint: '10' } } }]
        ];
        var tableStyle = { border: true, rtl: true, align: 'right' };
        docxFile.createTable(table, tableStyle);
        var fileName = "table.docx";
        var out = fs_1.createWriteStream(dir + fileName);
        docxFile.generate(out);
        out.on('close', function () {
            console.log('Create file of ' + fileName + 'in directory');
            done();
            return officeGen.Promise = global.Promise;
        });
    });
});
//# sourceMappingURL=officegenTest.js.map