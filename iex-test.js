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

    function getPrices(){
        var queryString = "https://api.iextrading.com/1.0/stock/market/batch?symbols=";
        var fieldString = "";
        for(ticker in Object.keys(stocks)){
            queryString += encodeURIComponent(ticker) + "%2C"; // comma
        }
        queryString.slice(0,-1); // remove last comma
        queryString += "&types=quote&filter=";
        for(field in Object.keys(fields)){
            fieldString += fieldMapping[field] + "%2C";
        }
        queryString.slice(0, -1);

        // call the service
        httpGetAsync(queryString, function(data){
            var result = JSON.parse(data);
            for(ticker in Object.keys(stocks)){
                for(field in Object.keys(stocks[ticker])){
                    for(callback in stocks[ticker][field]){
                        callback(result[ticker]["quote"][fieldMapping[field]]);
                    }
                }
            }
        });

        setTimeout(getPrices,500);
    }

    function quote (ticker, field, setResult) {
        // add the callback to memory
        if(!stocks[ticker]){
            stocks[ticker] = {};
        }    
        if(!stocks[ticker][field]){
            stocks[ticker][field] = [];
        }
        stocks[ticker][field].push(setResult);
        if(!fields[field]){
            fields[field] = true;
        }

        // start getting prices
        // assumes the ticker is valid
        if(!sessionActive){
            sessionActive = true;
            getPrices();
        }
    }   

    Excel.Script.CustomFunctions["STOCKS"]["QUOTE"] = {
        call: quote,
        description: "Get real-time market data from the IEX exchange.",
        helpUrl: "https://www.michael-saunders.com/help.html",
        result: {
            resultType: Excel.CustomFunctionValueType.number,
            resultDimensionality: Excel.CustomFunctionDimensionality.scalar,
        },
        parameters: [
            {
                name: "The stock ticker (eg. 'MSFT')",
                description: "The stock ticker (eg. 'MSFT')",
                valueType: Excel.CustomFunctionValueType.string,
                valueDimensionality: Excel.CustomFunctionDimensionality.scalar,
            },
            {
                name: "The field to query ('last', 'change', 'bid', 'ask', or 'volume')",
                description: "The field to query ('last', 'change', 'bid', 'ask', or 'volume')",
                valueType: Excel.CustomFunctionValueType.string,
                valueDimensionality: Excel.CustomFunctionDimensionality.scalar,
            }
        ],
        options:{ batch: false, stream: true }
    };
    
    

    // incrementValue is an example of a streaming function.
    function incrementValue(increment, setResult){    
    	var result = 0;
        setInterval(function(){
            result += increment;
            setResult(result);
        }, 1000);
    }
    
    
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