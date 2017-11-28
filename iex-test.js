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
    var socket = io('https://ws-api.iextrading.com/1.0/tops');
    socket.on('connect',function(){  });
    socket.on('message', function(message){
        var parsedMessage = JSON.parse(message);
        stocks[parsedMessage["symbol"]](parsedMessage["lastSalePrice"]);
    });

    
    var fields = {};

    function quote (ticker, field, invocationContext) {
        // add the callback to memory
        
        if(!stocks[ticker]){
            stocks[ticker] = invocationContext.setResult;
            socket.emit('subscribe', ticker);
        }
        
        httpGetAsync("https://api.iextrading.com/1.0/stock/" + ticker + "/quote?filter=iexRealtimePrice",function(data){
            var parsedData = JSON.parse(data);
            invocationContext.setResult(parsedData["iexRealtimePrice"]);
        });

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