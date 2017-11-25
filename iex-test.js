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