var officegen = require('officegen');
var docx = officegen({
    type: 'docx'
});
docx.on('finalize', function () {
    console.log('finish crate the file');
});
//# sourceMappingURL=officegen.js.map