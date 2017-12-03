var fieldMapping = {
    "last": "latestPrice",
    "bid": "iexBidPrice",
    "ask": "iexAskPrice",
    "volume": "iexVolume",
    "change": "change"
};

function httpGetAsync(theUrl, callback)
{
    var xmlHttp = new XMLHttpRequest();
    xmlHttp.onreadystatechange = function() { 
    if (xmlHttp.readyState == 4 && xmlHttp.status == 200)
        callback(xmlHttp.responseText);
    }
    xmlHttp.open("GET", theUrl, true);
    xmlHttp.send(null);
}

Office.initialize = function(reason){
    // Define the Contoso prefix.
    Excel.Script.CustomFunctions = {};
    Excel.Script.CustomFunctions["STOCKS"] = {};

    var stocks = {};
    var fields = {};
    var sessionActive = false;
    var counter = 0;

    function getPrices(){
        
        

        var queryString = "https://api.iextrading.com/1.0/stock/market/batch?symbols=";
        for(ticker in stocks){
            if(stocks.hasOwnProperty(ticker)){
                queryString += encodeURIComponent(ticker) + "%2C"; // comma
            }
        }
        queryString = queryString.slice(0,-3); // remove last comma
        queryString += "&types=quote&filter=";
        for(field in fields){
            if(fields.hasOwnProperty(field)){
                queryString += fieldMapping[field] + "%2C";
            }
        }
        queryString = queryString.slice(0, -3);

        // call the service
        httpGetAsync(queryString, function(data){
            var result = JSON.parse(data);
            for(ticker in stocks){
                if(stocks.hasOwnProperty(ticker)){
                    for(field in stocks[ticker]){
                        if(stocks[ticker].hasOwnProperty(field)){
                            for(var i = 0; i < stocks[ticker][field].length;i++){
                                stocks[ticker][field][i](result[ticker]["quote"][fieldMapping[field]]);
                            }
                        }
                    }
                }
            }
        });

        stocks["T"]["last"][0]("counter is " + counter);
        counter++;

        window.setTimeout(function(){
            getPrices();
        },500);
    }

    function quote (ticker, field, invocationContext) {
        // add the callback to memory
        
        if(!stocks[ticker]){
            stocks[ticker] = {};
        }    
        if(!stocks[ticker][field]){
            stocks[ticker][field] = [];
        }
        stocks[ticker][field].push(invocationContext.setResult);
        if(!fields[field]){
            fields[field] = true;
        }

        // start getting prices
        // assumes the ticker is valid
        if(!sessionActive){
            sessionActive = true;
            getPrices();
        }

        // remove entry if it's canceled
        invocationContext.onCanceled = function(){
            // remove the stock if there are no occurences 
        };

    }   

    Excel.Script.CustomFunctions["STOCKS"]["QUOTE"] = {
        call: quote,
        description: "Get real-time market data from the IEX exchange.",
        helpUrl: "https://www.michael-saunders.com/help.html",
        result: {
            resultType: Excel.CustomFunctionValueType.string,
            resultDimensionality: Excel.CustomFunctionDimensionality.scalar,
        },
        parameters: [
            {
                name: 'The stock ticker (eg. "MSFT")',
                description: "The stock ticker (eg. 'MSFT')",
                valueType: Excel.CustomFunctionValueType.string,
                valueDimensionality: Excel.CustomFunctionDimensionality.scalar,
            },
            {
                name: 'The field to query ("last", "change", "bid", "ask", or "volume")',
                description: "The field to query ('last', 'change', 'bid', 'ask', or 'volume')",
                valueType: Excel.CustomFunctionValueType.string,
                valueDimensionality: Excel.CustomFunctionDimensionality.scalar,
            }
        ],
        options:{ batch: false, stream: true }
    };
    
/*
    Excel.Script.CustomFunctions["STOCKS"]["DEBUG"] = {
        call: quote,
        description: "Debugging ",
        helpUrl: "https://www.michael-saunders.com/help.html",
        result: {
            resultType: Excel.CustomFunctionValueType.number,
            resultDimensionality: Excel.CustomFunctionDimensionality.scalar,
        },
        parameters: [
            {
                name: 'The stock ticker (eg. "MSFT")',
                description: "The stock ticker (eg. 'MSFT')",
                valueType: Excel.CustomFunctionValueType.string,
                valueDimensionality: Excel.CustomFunctionDimensionality.scalar,
            },
            {
                name: 'The field to query ("last", "change", "bid", "ask", or "volume")',
                description: "The field to query ('last', 'change', 'bid', 'ask', or 'volume')",
                valueType: Excel.CustomFunctionValueType.string,
                valueDimensionality: Excel.CustomFunctionDimensionality.scalar,
            }
        ],
        options:{ batch: false, stream: true }
    };
*/

    Excel.Promise = function (setResultFunction){
        return new OfficeExtension.Promise(function(resolve, reject){
            setResultFunction(resolve, reject);
        });
    }  
    
    // helper code for getting temperature
    var temps = {};
    temps["boiler"] = 104.3;
    temps["mixer"] = 44.0;
    temps["furnace"] = 586.9;
    furnaceHistory = [];
    function startTime(){
        temps["boiler"] += Math.pow(Math.random() - 0.45, 3) * 2;
        temps["mixer"] += Math.pow(Math.random() - 0.55, 3) * 2;
        temps["furnace"] += Math.pow(Math.random() - 0.40, 3) * 2;
        furnaceHistory.push([temps["furnace"]]);
        if(furnaceHistory.length > 50){
            furnaceHistory.shift();
        }
        setTimeout(startTime, 500);
    }
    startTime();
    function getTempFromServer(thermometerID, callback){
        setTimeout(function(){
            var data = {};
            data.temperature = temps[thermometerID].toFixed(1);
            callback(data);
        }, 200);
    }

    // demo functions

    function addTo42(num){
        return Excel.Promise(function(setResult, setError){
            setTimeout(function(){
                setResult(num + 42);
            }, 1000);
        });
    }
    
    function addTo42Fast(num) {
        return num + 42;
    }

    function getTemperature(thermometerID){ 
        return Excel.Promise(function(setResult, setError){ 
            getTempFromServer(thermometerID, function(data){ 
                setResult(data.temperature); 
            }); 
        }); 
    }

    function streamTemperature(thermometerID, interval, call){     
        if(thermometerID == "furnace"){
            temps["furnace"] = 630.2;
        }
        function getNextTemperature(){ 
            getTempFromServer(thermometerID, function(data){ 
                call.setResult(data.temperature); 
            }); 
            setTimeout(getNextTemperature, interval); 
        } 
        getNextTemperature(); 
    } 

    function secondHighestTemp(temperatures){ 
        var highest = -273, secondHighest = -273;
        for(var i = 0; i < temperatures.length;i++){
            for(var j = 0; j < temperatures[i].length;j++){
                if(temperatures[i][j] >= highest){
                    secondHighest = highest;
                    highest = temperatures[i][j];
                }
                else if(temperatures[i][j] >= secondHighest){
                    secondHighest = temperatures[i][j];
                }
            }
        }
        return secondHighest;
    }

    function trackTemperature(thermometerID, call){
        var output = [];
        
        for(var i = 0; i < 50; i++) output.push([0]);  
        if(thermometerID == "furnace"){
            output = furnaceHistory;
        } 
        function recordNextTemperature(){
            getTempFromServer(thermometerID, function(data){
                output.push([data.temperature]);
                output.shift();
                call.setResult(output);
            });
            setTimeout(recordNextTemperature, 500);
        }
        recordNextTemperature();
    } 

    Excel.Script.CustomFunctions["CFACTORY"] = {};
    Excel.Script.CustomFunctions["CFACTORY"]["ADDTO42"] = {
        call: addTo42Fast,
        description: "Returns the sum of a number and 42, fast",
        helpUrl: "https://example.com/help.html",
        result: {
            resultType: Excel.CustomFunctionValueType.number,
            resultDimensionality: Excel.CustomFunctionDimensionality.scalar,
        },
        parameters: [
            {
                name: "num",
                description: "The number be added",
                valueType: Excel.CustomFunctionValueType.number,
                valueDimensionality: Excel.CustomFunctionDimensionality.scalar,
            },
        ],
        options: {
            batch: false,
            stream: false,
        }
    };
    Excel.Script.CustomFunctions["CFACTORY"]["GETTEMPERATURE"] = {
        call: getTemperature,
        description: "Returns the temperature of the boiler, mixer, or furnace",
        helpUrl: "https://example.com/help.html",
        result: {
            resultType: Excel.CustomFunctionValueType.number,
            resultDimensionality: Excel.CustomFunctionDimensionality.scalar,
        },
        parameters: [
            {
                name: "thermometer ID",
                description: "The thermometer to be measured",
                valueType: Excel.CustomFunctionValueType.string,
                valueDimensionality: Excel.CustomFunctionDimensionality.scalar,
            },
        ],
        options: {
            batch: false,
            stream: false,
        }
    };
    Excel.Script.CustomFunctions["CFACTORY"]["STREAMTEMPERATURE"] = {
        call: streamTemperature,
        description: "Streams the temperature of the boiler, mixer, or furnace",
        helpUrl: "https://example.com/help.html",
        result: {
            resultType: Excel.CustomFunctionValueType.number,
            resultDimensionality: Excel.CustomFunctionDimensionality.scalar,
        },
        parameters: [
            {
                name: "thermometer ID",
                description: "The thermometer to be measured",
                valueType: Excel.CustomFunctionValueType.string,
                valueDimensionality: Excel.CustomFunctionDimensionality.scalar,
            },
            {
                name: "interval",
                description: "The time between updates",
                valueType: Excel.CustomFunctionValueType.number,
                valueDimensionality: Excel.CustomFunctionDimensionality.scalar,
            },
        ],
        options: {
            batch: false,
            stream: true,
        }
    };
    Excel.Script.CustomFunctions["CFACTORY"]["SECONDHIGHESTTEMP"] = {
        call: secondHighestTemp,
        description: "Returns the second highest from a range of temperatures",
        helpUrl: "https://example.com/help.html",
        result: {
            resultType: Excel.CustomFunctionValueType.number,
            resultDimensionality: Excel.CustomFunctionDimensionality.scalar,
        },
        parameters: [
            {
                name: "temps",
                description: "the temperatures to be compared",
                valueType: Excel.CustomFunctionValueType.number,
                valueDimensionality: Excel.CustomFunctionDimensionality.matrix,
            },
        ],
        options: {
            batch: false,
            stream: false,
        }
    };
    Excel.Script.CustomFunctions["CFACTORY"]["TRACKTEMPERATURE"] = {
        call: trackTemperature,
        description: "Streams 25 seconds of temperature history",
        helpUrl: "https://example.com/help.html",
        result: {
            resultType: Excel.CustomFunctionValueType.number,
            resultDimensionality: Excel.CustomFunctionDimensionality.matrix,
        },
        parameters: [
            {
                name: "thermometer ID",
                description: "The thermometer to be measured",
                valueType: Excel.CustomFunctionValueType.string,
                valueDimensionality: Excel.CustomFunctionDimensionality.scalar,
            },
        ],
        options: {
            batch: false,
            stream: true,
        }
    };

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