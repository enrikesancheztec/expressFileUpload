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

app.post("/api/movimientos", upload.single('file'), uploadMovimientos);
function uploadMovimientos(req, res) {    
    var workbook = xlsx.read(req.file.buffer);
    var sheet_name_list = workbook.SheetNames;
    var data = xlsx.utils.sheet_to_json(workbook.Sheets[sheet_name_list[0]]);
    var account = "";
    var output = "";
    var count = 0;
    for (let line of data) {
        if (Object.keys(line).length >= 4 && line["CONTPAQ i"] !== undefined && line["__EMPTY"] !== "") {
            if (String(line["CONTPAQ i"]).match(/\d{3}-\d{3}/)) {
                account = line["CONTPAQ i"];
            } else {
                if (Object.keys(line).length >= 6 && line["CONTPAQ i"] !== "Fecha") {
                    output += "REGISTRO " + (++count) + "-----------------------------\n";
                    output += "CUENTA: " + account + "\n";  
                    output += "FECHA: " + line["CONTPAQ i"] + "\n";  
                    output += "TIPO: " + line["__EMPTY"] + "\n";  
                    output += "NUMERO: " + line["__EMPTY_1"] + "\n";
                    output += "CONCEPTO: " + line["Lecar Consultoria en TI, S.C."] + "\n";
                    output += "REFERENCIA: " + line["__EMPTY_2"] + "\n";
                    output += "CARGOS: " + line["__EMPTY_3"] + "\n";
                    output += "ABONOS: " + line["__EMPTY_4"] + "\n";
                    output += "SALDO: " + line["Hoja:      1"] + "\n";  
                }
            }
        }
    }
    console.log(output);

    return res.status(201).send(output);
}

app.listen(5000, () => {
    console.log(`Server started...`);
});