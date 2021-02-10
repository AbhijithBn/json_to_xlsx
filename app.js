const https = require('https');
var express = require("express");
var cors = require('cors')
var app = express();

const xlsx = require('xlsx');
const path = require('path');
const fs = require('fs');

// const options = {
//     key: fs.readFileSync('key.pem'),
//     cert: fs.readFileSync('cert.pem')
// };


//server port
var PORT=process.env.PORT||8000;

//To avoid Cross Origin errors
app.use(cors());

//body parser
var bodyParser = require("body-parser");
app.use(bodyParser.json());
app.use(bodyParser.urlencoded({ extended: true }));

//Array of JSON post requests
var detailsArray = [];

function createNewFile (){
    if(detailsArray.length!=0){
        const excelHeader = [
            "Event",
            "FlowID",
            "User",
            "Source"
        ]
        var details = detailsArray;
        const excelSheetName = 'FlowDetails';

        let d = new Date();
        let hour = d.getHours();
        let min = d.getMinutes();
        let seconds = d.getSeconds();
        const filePath = "../"+hour+min+seconds+".xlsx";

        function exportExcel(data, excelHeader, excelSheetName, filePath){
            const workBook = xlsx.utils.book_new();
            const workSheetData = [
                excelHeader,
                ... data
            ];
            const workSheet = xlsx.utils.aoa_to_sheet(workSheetData);
            xlsx.utils.book_append_sheet(workBook, workSheet, excelSheetName);
            xlsx.writeFile(workBook, path.resolve(filePath));
        }
        
        function exportUsersToExcel(details, excelHeader, excelSheetName, filePath){
            const data = details.map(detail => {
                return [detail.event, detail.flowid, detail.user,detail.source];
            });
            console.log("Data is :",data," type is ",typeof(data));
            exportExcel(data, excelHeader, excelSheetName, filePath);
        }
        exportUsersToExcel(details, excelHeader, excelSheetName, filePath);

        detailsArray = [];
    }
}

let checkInterval = setInterval(createNewFile, 60000);

app.get('/details',(req,res)=>{
    res.send(detailsArray);
})

//POST request
app.post('/data',(req,res)=>{
    detailsArray.push(req.body);
    console.log(detailsArray.length);
    if(detailsArray.length>=5){
        createNewFile();
    }
    res.send("Post Request Called");
})

var server=app.listen(PORT,function(){
    var host=server.address().address
    var port=server.address().port
    console.log("Example app listening at http://%s:%s", host, port)
})

// https.createServer(options, app).listen(443);