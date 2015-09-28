# json2xl
Exports JSON data to Excel along with Styles, Formatting and Formulas
 
## Usage
 
    var json2xl = require("json2xl");
    
    app.get("/excelexport", function (req, res) {
      var filepath = "uploads/"
      var fileName = Date.now() + '.xlsx';
      
      var wbOpts = {
            jszip:{
                compression:'DEFLATE'
            }
        };
 
      var wsOpts = {
            margins:{
                left : .75,
                right : .75,
                top : 1.0,
                bottom : 1.0,
                footer : .5,
                header : .5
            },
            printOptions:{
                centerHorizontal : true,
                centerVertical : false
            },
            view:{
                zoom : 100
            },
            outline:{
                summaryBelow : true
            },
            fitToPage:{
                fitToHeight: 100,
                orientation: 'landscape',
          },
        }
 
      var data = {
              "worksheets" : ['Page-1'],                 
              "filepath": filepath,
              "filename": fileName,                  
              "rows": [
                    ['Serial','Age',"Name"],
                    [1,27,"Test-1"],
                    [2,27,"Test-2"],
                    [3,27,"Test-3"],
                    [4,27,"Test-4"],
                    [5,27,"Test-5"],
                    [6,27,"Test-6"]
                ],
                "rows" : [
                             [
                                {   
                                    "value" : "Row-1-Col-1",
                                    "dataType": "string",
                                    "style":[{
                                        "color" : "red",
                                        "backgroundColor" : "green",
                                        "border" : ["thick", "black"]
                                    }]                                       
                                },
                                {   
                                    "value" : "2", 
                                    "dataType": "number",
                                    "style":[{
                                        "color" : "green",
                                        "backgroundColor" : "brown",
                                        "border" : ["thick","blue"]
                                    }]
                                },
                                {   
                                    "value" : "3", 
                                    "dataType": "number",
                                    "style":[{
                                        "color" : "green",
                                        "backgroundColor" : "red"
                                    }]
                                },
                                {   
                                    "value" : 'B1+C1', 
                                    "dataType": "Formula",
                                    "style":[{
                                        "color" : "green",
                                        "backgroundColor" : "red"
                                    }]
                                },
                             ],
                           ],                    
                  "config" : { 
                        "wbOpts" : wbOpts,
                        "wsOpts" : wsOpts,
                        "freezePanes" : {
                            "rows" : [1],
                            "cols" : [3]
                        }
                    }
              };
             
             Utility.ExportToExcelWithStyles(data, function(err, response){
                res.end(response);
             });
 
Prose