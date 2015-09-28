## [[Json2xl]()](https://github.com/hansiemithun/json2xl "Json2XL")
Exports JSON data to Excel along with Styles, Formatting and Formulas
 
### [Usage]()
 
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
 
### [Datatypes: (Optional)]()
  1. String 
  2. Number
  3. Formula
  4. Link
  5. Date

All the datatypes are optional. If nothing is defined "String" dataType is considered

### [CSS Styles]()
  1. color
  2. fontSize
  3. backgroundColor
  
 "color" & "backgroundColor" can be either in hexadecimal or color names : #F00 or Red
 "fontSize" should be in pixels : 11px
 
 ### [Default CSS Styles]()
 1. Pattern: "Solid"
 2. Color: "#000" or "Black"
 3. BackgroundColor: "#FFF" or "White"
 4. FontSize: 10px
 5. FontFamily: "Arial"
 6. FontWeight: "Normal"
 7. BorderColor: "black"


 
  
  
  




