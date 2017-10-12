let officegen = require('officegen');

let docx  =  officegen ({
    type: 'docx'
});


docx.on('finalize', function(){
    console.log('finish crate the file');
});

//console.log("heloo");