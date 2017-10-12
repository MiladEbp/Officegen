import {} from 'mocha';
import {createWriteStream} from 'fs';

let officeGen = require('officegen');
let dir       = "outStream/";

describe('officegen', function(){

    it('Add Header in File',  function(done){
        let docxFile = officeGen('docx');
        let header   = docxFile.getHeader().createP({align:'center'});
        // Add Text For Header
        header.addText('Unit Test For officegen', {color: 'black'});
        //create file
        let fileName = 'header.docx';
        let out      = createWriteStream(dir+fileName);
            docxFile.generate(out);
            out.on('close',function(){
                console.log('OutPut is : \nCreate file of '+fileName+' in directory');
                done();
                return officeGen.Promise = global.Promise;
            });// out.On

    });// it for Add Header


    it('Add Description in File', function (done){
        let docxFile = officeGen('docx');
        let createP  = docxFile.createP({align: 'center' });
        // Add Text //
        createP.addText (  'این فایل برای تست است .' ,{color: 'red', font_face: 'B Nazanin'});
            // Create File //
            let fileName = 'description.docx';
            let out      = createWriteStream(dir+fileName);
                docxFile.generate(out);
                out.on('close', function(){
                    console.log('Create file of '+fileName+' in directory');
                     done();
                     return officeGen.Promise = global.Promise;
                });// out.on

    });// it for Add Description


    it('Add Table in File', function(done){
        let docxFile   = officeGen('docx');
        let table      = [ [{val: 'نام' , opts:{rtl: true ,font_face: 'B Nazanin' ,color:'block',shd:{fill: '7F7F7F', themeFill: 'text1' ,themeFillTint: '10'}}},
                            {val: 'نام خانوادگی' , opts:{rtl: true ,font_face: 'B Nazanin' ,color:'block',shd:{fill: '7F7F7F', themeFill: 'text1' ,themeFillTint: '80'}}}] ,
                            [{val: 'عرفان', opts:{rtl: true ,font_face: 'B Nazanin' ,color:'red'  ,shd:{fill: '92CDDC', themeFill: 'text1' ,themeFillTint: '80'}}},
                             {val: 'تاواتاو', opts:{rtl: true ,font_face: 'B Nazanin' ,color:'red'  ,shd:{fill: '92CDDC', themeFill: 'text1' ,themeFillTint: '10'}}}]
                         ];// table

        let tableStyle = {border: true , rtl:true , align: 'right'};

        // Create Table
           docxFile.createTable(table , tableStyle);
           let fileName = "table.docx";
           let out      = createWriteStream(dir+fileName);

           docxFile.generate(out);
                 out.on('close', function(){
                     console.log('Create file of '+fileName+'in directory');
                     done();
                     return officeGen.Promise = global.Promise;
                 });// out.on
    });// it For Add Table in File


});// describe