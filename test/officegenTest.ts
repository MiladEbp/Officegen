import {} from 'mocha';
import {createWriteStream} from 'fs';


let officeGen = require('officegen');
let dir = "outStream/";

describe('officegen', function(){

    it('Add Header in File',  function(done){
        let docxFile = officeGen('docx');
        let header   = docxFile.getHeader().createP({align:'center'});
        // Add Text For Header
        header.addText('Unit Test For officegen', {color: 'black'});
        //create file
        let fileName = 'header.docx';
        let out = createWriteStream(dir+fileName);
            docxFile.generate(out);
            out.on('close',function(){
                console.log('Create file of '+fileName+' in directory');
                done();
                return officeGen.Promise = global.Promise;
            });// out.On

    });// it for Add Header


    it('Add Description in File', function (done){
        let docxFile = officeGen('docx');
        let createP = docxFile.createP({align: 'center' });
        // Add Text //
        createP.addText (  'این فایل برای تست است .' ,{color: 'red', font_face: 'B Nazanin'});
            // Create File //
            let fileName = 'description.docx';
            let out = createWriteStream(dir+fileName);
                docxFile.generate(out);
                out.on('close', function(){
                    console.log('Create file of '+fileName+' in directory');
                     done();
                     return officeGen.Promise = global.Promise;
                });// out.on

    });// it for Add Description



});// describe