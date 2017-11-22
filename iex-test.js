/* 
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

Office.initialize = function(reason){
    // Define the Contoso prefix.
    Excel.Script.CustomFunctions = {};
    Excel.Script.CustomFunctions["CONTOSO"] = {};

    // add42 is an example of a synchronous function.
    function add42 (a, b) {
        return a + b + 42;
    }    
    Excel.Script.CustomFunctions["STOCKS"]["ADD42"] = {
        call: add42,
        description: "Finds the sum of two numbers and 42.",
        helpUrl: "https://www.contoso.com/help.html",
        result: {
            resultType: Excel.CustomFunctionValueType.number,
            resultDimensionality: Excel.CustomFunctionDimensionality.scalar,
        },
        parameters: [
            {
                name: "num 1",
                description: "The first number",
                valueType: Excel.CustomFunctionValueType.number,
                valueDimensionality: Excel.CustomFunctionDimensionality.scalar,
            },
            {
                name: "num 2",
                description: "The second number",
                valueType: Excel.CustomFunctionValueType.number,
                valueDimensionality: Excel.CustomFunctionDimensionality.scalar,
            }
        ],
        options:{ batch: false, stream: false }
    };
    
    

    // incrementValue is an example of a streaming function.
    function incrementValue(increment, setResult){    
    	var result = 0;
        setInterval(function(){
            result += increment;
            setResult(result);
        }, 1000);
    }
    Excel.Script.CustomFunctions["STOCKS"]["INCREMENTVALUE"] = {
        call: incrementValue,
        description: "Increments a counter that starts at zero.",
        helpUrl: "https://www.contoso.com/help.html",
        result: {
            resultType: Excel.CustomFunctionValueType.number,
            resultDimensionality: Excel.CustomFunctionDimensionality.scalar,
        },
        parameters: [
            {
                name: "period",
                description: "The time between updates, in milliseconds.",
                valueType: Excel.CustomFunctionValueType.number,
                valueDimensionality: Excel.CustomFunctionDimensionality.scalar,
            },
        ],
        options: { batch: false,  stream: true }
    };
    
    // The refreshTemperature and streamTemperature functions use global variables to save & read state, while streaming data.
    var savedTemperatures = {};
    function refreshTemperature(thermometerID){        
        sendWebRequestExample(thermometerID, function(data){
            savedTemperatures[thermometerID] = data.temperature;
        });
        setTimeout(function(){
            refreshTemperature(thermometerID);
        }, 1000);
    }
    function streamTemperature(thermometerID, setResult){    
        if(!savedTemperatures[thermometerID]){
            refreshTemperature(thermometerID);
        }
        function getNextTemperature(){
            setResult(savedTemperatures[thermometerID]);
            setTimeout(getNextTemperature, 1000);
        }
        getNextTemperature();
    }

    // Register all the custom functions previously defined in Excel.
    Excel.run(function (context) {        
        context.workbook.customFunctions.addAll();
        return context.sync().then(function(){});
    }).catch(function(error){});

    // The following are helper functions.

    // The log function lets you write debugging messages into Excel (first evaluate the MY.DEBUG function in Excel). You can also debug with regular debugging tools like Visual Studio.
    var debug = [];
    var debugUpdate = function(data){};
    function log(myText){
        debug.push([myText]);
        debugUpdate(debug);
    }
    function myDebug(setResult){
        debugUpdate = setResult;
    }
   
}; 