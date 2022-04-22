const express = require("express");
const multer = require("multer");
const xlsx = require('xlsx');

const app = express();
app.use(express.json());
app.use(express.urlencoded({ extended: true }));

var storage = multer.memoryStorage();

var upload = multer({
    storage: storage
});

app.post("/api/xlsx", upload.single('file'), uploadXlsx);
function uploadXlsx(req, res) {    
    var workbook = xlsx.read(req.file.buffer);
    var sheet_name_list = workbook.SheetNames;
    var data = xlsx.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]);
    console.log(data);
    return res.status(201).send(data);
}

app.listen(5000, () => {
    console.log(`Server started...`);
});