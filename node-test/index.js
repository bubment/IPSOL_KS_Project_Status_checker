var Firebird = require('node-firebird');
const async = require('async');
const fs = require('fs');
var xl = require('excel4node');
const winston = require('winston');
let spawn = require("child_process").spawn,child;

const logger = winston.createLogger({
    level: 'info',
    format: winston.format.json(),
    defaultMeta: { service: 'user-service' },
    transports: [
      //
      // - Write all logs with level `error` and below to `error.log`
      // - Write all logs with level `info` and below to `combined.log`
      //
      new winston.transports.File({ filename: 'error.log', level: 'error' }),
      new winston.transports.File({ filename: 'combined.log' }),
    ],
  });
   
  //
  // If we're not in production then log to the `console` with the format:
  // `${info.level}: ${info.message} JSON.stringify({ ...rest }) `
  //
  if (process.env.NODE_ENV !== 'production') {
    logger.add(new winston.transports.Console({
      format: winston.format.simple(),
    }));
  }

var options = {};
 
options.host = 'localhost';
options.port = 3056;
options.database = 'C:/ProgramData/KS/FbDatabaseServer/Databases/KSCOMPANY_IPSOLRENDSZERHZZRT_191114155447.KSFDB';
options.user = 'bubment';
options.password = 'Ips0l2020';
options.lowercase_keys = false; // set to true to lowercase keys
options.role = null;            // default
options.pageSize = 12000;        // default when creating database


function mainFunction(){

    //Result array of Firebird SQL queries
    let reportData = [
        {type: 'Voucher', data:[]},
        {type: 'Project', data:[]}
    ];

    //Firebird SQL query infos.
    let queryInfo = [
        {type: 'Voucher', query:'SELECT vs."Caption" AS "VoucherSequence", cso."VoucherNumber", cso."VoucherDate", cso."CustomerNameDisplay", cso."NetValue", c."Name" AS "Currency" , cso."CurrencyRate", d."Name" AS "Division" ,b."Subject" AS "DiscountName", CAST(csod."DetailComment" AS varchar(8000)) AS "DetailComment" FROM "CustomerStockOut" cso LEFT JOIN "VoucherSequence" vs ON cso."VoucherSequence" = vs."Id" LEFT JOIN "Business" b ON cso."Business" = b."Id" LEFT JOIN "Currency" c ON cso."Currency" = c."Id" LEFT JOIN "Division" d ON CSO."Division" = d."Id" LEFT JOIN "CustomerStockOutDetail" csod ON cso."Id" = csod."CustomerStockOut"'},
        {type: 'Voucher', query:'SELECT vs."Caption" AS "VoucherSequence", co."VoucherNumber", co."VoucherDate" , c2."Name" AS "CustomerNameDisplay", co."NetValue", c."Name" AS "Currency" , co."CurrencyRate", d."Name" AS "Division" , b."Subject" AS "DiscountName", CAST(cod."DetailComment" AS varchar(8000)) AS "DetailComment"  FROM "CustomerOrder" co LEFT JOIN "VoucherSequence" vs ON co."VoucherSequence" = vs."Id" LEFT JOIN "Business" b ON co."Business" = b."Id" LEFT JOIN "Currency" c ON co."Currency" = c."Id" LEFT JOIN "Division" d ON co."Division" = d."Id" LEFT JOIN "CustomerOrderDetail" cod ON co."Id" = cod."CustomerOrder" LEFT JOIN "Customer" c2 ON co."Customer" = c2."Id"'},
        {type: "Project", query:'SELECT b."Subject", c."Name" AS "Customer", bs."Name" AS "BusinessState", u."Name" AS "User", ct."Name" AS "JobNumber" FROM "Business" b LEFT JOIN "Customer" c ON b."Customer" = c."Id" LEFT JOIN "BusinessState" bs ON b."BusinessState" = bs."Id" LEFT JOIN "User" u ON b."User" = u."Id" LEFT JOIN CT_133777 ct ON b."CF_133778" = ct."Id"'}
    ]

    let copyDBFile = function(callback){

        fs.copyFile('C:/KS_fajlok/node-test/db/security/security2.fdb','C:/ProgramData/KS/FbDatabaseServer/security2.fdb',(err)=>{
                if (err) {
                    if (err.code == 'EBUSY') {
                        callback();
                    }else{
                        console.log('There was an issue copying the security2.fdb file.')
                        callback(err);
                    }
                    
                }else{
                    console.log("security2.fdb file copied successfully");
                    callback();
                }
            });
    }

    let firebirdQueries = function(callback){

        let queryFunction = function(item, callback){
            Firebird.attach(options, function(err, db) {
                if (err){
                    console.log("There was a problem to connect to the database.");
                    callback(err);
                }
                else{
                    db.query(item.query, function(err, result) {
                        if (err) {
                            db.detach();
                            console.log("There was a problem with the following SQL query:")
                            console.log(item.query);
                            callback(err)
                        }else{
                            let actType;
                            for (let i = 0; i < reportData.length; i++) {
                                if (reportData[i].type == item.type) {
                                    actType = i;
                                    break;
                                }
                            }
                            if (actType != undefined) {
                                for (let i = 0; i < result.length; i++) {
                                    if (item.type == "Voucher") {
                                        result[i].NetValueHUF = result[i].NetValue * result[i].CurrencyRate
                                    }
                                    reportData[actType].data.push(result[i])
                                }
                                //Duplikációszűrés
                                for (let i = 0; i < reportData[actType].data.length; i++) {
                                    for (let j = (i+1); j < reportData[actType].data.length; j++){
                                        if (reportData[actType].data[i] == reportData[actType].data[j]) {
                                            reportData[actType].splice(j,1);
                                            j--;
                                        }       
                                    }
                                }

                                db.detach();
                                callback();
                            }else{
                                console.log("There was a problem on firebirdQueries.")
                                db.detach();
                                callback({message: "reportData array has not got any item with type " + item.type})
                            }
                        }
                    });
                }   
            });
        }

        async.eachSeries(
            queryInfo,
            queryFunction,
            function(err){
                if(err){
                    callback(err);
                }else{
                    console.log("firebirdQueries finished successfully");
                    callback();
                }
            }
        )
    }

    let createExcel = function(callback){

        let excelInfo = [
            {
                fileName : "Bizonylatok tételes lekérdezése.xlsx",
                type: "Voucher",
                columnInfo : [
                        {ColumnHU: 'Bizonylat fajta', ColumnEN:'VoucherSequence', ColumnNumber:1},
                        {ColumnHU: 'Bizonylatszám', ColumnEN:'VoucherNumber', ColumnNumber:2},
                        {ColumnHU: 'Kelte', ColumnEN:'VoucherDate', ColumnNumber:3},
                        {ColumnHU: 'Ügyfél', ColumnEN:'CustomerNameDisplay', ColumnNumber:7},
                        {ColumnHU: 'Egységár', ColumnEN:'NetValue', ColumnNumber:18},
                        {ColumnHU: 'Pénznem', ColumnEN:'Currency', ColumnNumber:20},
                        {ColumnHU: 'Árfolyam', ColumnEN:'CurrencyRate', ColumnNumber:21},
                        {ColumnHU: 'Nettó érték (HUF)', ColumnEN:'NetValueHUF', ColumnNumber:26},
                        {ColumnHU: 'Részlegszám', ColumnEN:'Division', ColumnNumber:29},
                        {ColumnHU: 'Projekt', ColumnEN:'DiscountName', ColumnNumber:30},
                        {ColumnHU: 'Megjegyzés', ColumnEN:'DetailComment', ColumnNumber:31}
                    ]
            },
            {
                fileName : "Projektek listája.xlsx",
                type: "Project",
                columnInfo : [
                        {ColumnHU: 'Tárgy', ColumnEN:'Subject', ColumnNumber:1},
                        {ColumnHU: 'Ügyfél név', ColumnEN:'Customer', ColumnNumber:2},
                        {ColumnHU: 'Állapot', ColumnEN:'BusinessState', ColumnNumber:5},
                        {ColumnHU: 'Kapcsolattartó', ColumnEN:'User', ColumnNumber:6},
                        {ColumnHU: "'Projekt kategória'", ColumnEN:'JobNumber', ColumnNumber:7},
                    ]
            }
        ]

        let writeExcelFile = function(item,callback){
            var wb = new xl.Workbook();
 
            // Add Worksheets to the workbook
            var ws = wb.addWorksheet('Sheet');

            //Put the Heared values to the worksheet
            for (let i = 0; i < item.columnInfo.length; i++) {
                ws.cell(1,item.columnInfo[i].ColumnNumber).string(item.columnInfo[i].ColumnHU);
            }

            let actType;
            for (let i = 0; i < reportData.length; i++) {
                if (reportData[i].type == item.type) {
                    actType = i;
                    break;
                }
            }

            //Put the content to the worksheet
            let actValue;
            for (let i = 0; i < reportData[actType].data.length; i++) {
                for (let j = 0; j < item.columnInfo.length; j++) {
                    if (reportData[actType].data[i][item.columnInfo[j].ColumnEN]) {
                        actValue = reportData[actType].data[i][item.columnInfo[j].ColumnEN]
                        
                        //Excel makró miatti javítás
                        if (actValue == "Teljesítési Igazolás") {
                            actValue = "Vevői rendelés"
                        }
                        
                        switch (typeof actValue) {
                            case "number":
                                ws.cell(i+2,item.columnInfo[j].ColumnNumber).number(actValue)
                                break;
                            case "string":
                                ws.cell(i+2,item.columnInfo[j].ColumnNumber).string(actValue)
                                break;
                            case "boolean":
                                ws.cell(i+2,item.columnInfo[j].ColumnNumber).bool(actValue)
                                break;
                            default:
                                ws.cell(i+2,item.columnInfo[j].ColumnNumber).string(actValue.toString());
                                break;
                        }
                    }
                }
            }

            wb.write('output/' + item.fileName, function(err) {
                if (err) {
                    console.log("There was an error durring writeExcelFile function.")
                    callback(err)
                } else {
                    console.log("writeExcelFile function finished successfully.")
                    callback()
                }
              });
        }

        async.eachSeries(
            excelInfo,
            writeExcelFile,
            function(err){
                if (err) {
                    console.log(err)
                    console.log("There was a problem with the createExcel function");
                    callback(err);
                }else{
                    console.log("createExcel finished successfully.")
                    callback();
                }
            }
        )
    }

    let runShellScripts = function(callback){
        let shellParams = [
            {
                psFile:"uploadToSharePoint",
                params: [{name:'moveableFile', value: '"C:/KS_fajlok/node-test/output/Bizonylatok tételes lekérdezése.xlsx"'}]
            },
            {
                psFile:"uploadToSharePoint",
                params: [{name:'moveableFile', value: '"C:/KS_fajlok/node-test/output/Projektek listája.xlsx"'}]
            }
        ]

        let shellFunction = function(item,callback){
            let callbackHandled = false;
            let parameterText = "";
            for (let i = 0; i < item.params.length; i++) {
                parameterText+= " -" + item.params[i].name + " " + item.params[i].value
            }

            child = spawn("powershell.exe",["./script/shell/" + item.psFile + ".ps1" + parameterText]);
            child.stdout.setEncoding('utf-8');

            child.stderr.on("data",function(data){
                if (!callbackHandled) {
                    console.log("Powershell Errors: " + data);
                    callbackHandled = true;
                    callback({message: data});
                }

            });
            child.on("exit",function(){
                console.log(item.psFile, " script finished");
                if (!callbackHandled) {
                    callback();
                }
            });
            child.stdin.end(); //end input 
        }

        async.eachSeries(
            shellParams,
            shellFunction,
            function(err){
                if (err) {
                    console.log(err)
                    console.log("There was a problem with the runShellScripts function");
                    callback(err);
                }else{
                    console.log("runShellScripts finished successfully.")
                    callback();
                }
            }
        )
    }
    
    async.series([
        copyDBFile,
        firebirdQueries,
        createExcel,
        runShellScripts
        

    ],function(err){
        if (err) {
            console.log("There was an error durring mainFunction")
            console.log('--------------------------------------------------------------------')
            console.log('--------------------------------------------------------------------')
            console.error("Error: ", err.message);
            console.log('--------------------------------------------------------------------')
            console.log('--------------------------------------------------------------------')
        }else{
            logger.info({'UTC':new Date()});
            console.log("mainFunction finished successfully.")
        }
    })
}

mainFunction()

