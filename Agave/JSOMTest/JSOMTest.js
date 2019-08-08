// Originated from
// src\osftest\files\MOEs\JSOMSettings.js, \\daddev\Office\16.0\19104.12000\src\osftest\files\MOEs
// JSOMEvents_ExcelIncontent.js

var result;

// ========== Variables: Binding ==========
var _bindingType;
var _bindingObj;


// ========== Variables: Setting ==========
var _Bindings;
var _Settings;
var names = ['String', 'Status', 'Integer', 'Array of integers', 'key_update', 'key_delete'];
var values = [];
values[0] = 'This is a string value'; //string value
values[1] = true; //boolean value
values[2] = 45; //number
values[3] = [5, 8, 10, 100]; //array
values[4] = 'val1'; // AgaveSettingUpdate_Succeed, CrossXlApiOsfTest::Coauth_AgaveSettingUpdate_Success
values[5] = 'val1'; // AgaveSettingDelete_Succeed, CrossXlApiOsfTest::Coauth_AgaveSettingDelete_Success


// ========== Init ==========
Office.initialize = function (reason)
{
    _Settings = Office.context.document.settings;
    _Bindings = Office.context.document.bindings;
    if (undefined === _Settings) 
    {
        return;
    }
}


// ========== Functions: Binding ==========
function getBindingId()
{
    return document.getElementById("bindingId").value;
}

function addBinding()
{
    var bindingId = getBindingId();
    if (!bindingId || bindingId.trim().length == 0)
    {
        output("Error: Need to specify new binding id");
        return;
    }

    if (!_bindingType)
    {
        output("Error: Need to specify binding type");
        return;
    }

    _Bindings.addFromSelectionAsync(
        _bindingType,
        {
            id: getBindingId()
        },
        function (asyncResult)
        {
            var statusString = "Status: " + asyncResult.status;
            if (asyncResult.status == 'failed')
            {
                var _errorString = statusString + ", Name: " + asyncResult.error.name + ", Message: " + asyncResult.error.message;
                returnValueToAutomation(_errorString);
                _bindingObj = null;
            }
            else
            {
                _bindingObj = asyncResult.value;
                returnValueToAutomation(statusString + ", Id: " + _bindingObj.id + ", Type: " + _bindingObj.type);
            }
        }
    );
}

function removeBinding()
{
    var bindingId = getBindingId();
    if (!bindingId || bindingId.trim().length == 0)
    {
        output("Error: Need to specify binding id to remove");
        return;
    }

    _bindingObj = null;
    _Bindings.releaseByIdAsync(
        bindingId,
        function (asyncResult)
        {
            var statusString = "Status: " + asyncResult.status;
            if (asyncResult.status == 'failed')
            {
                var _errorString = statusString + ", Name: " + asyncResult.error.name + ", Message: " + asyncResult.error.message;
                returnValueToAutomation(_errorString);
            }
            else
            {
                returnValueToAutomation(statusString + ", Removed binding '" + bindingId + "'");
            }
        }
    );
}

function getBindingsCount()
{
    Excel.run(function(ctx) 
    {
        var bindings = ctx.workbook.bindings;
        bindings.load('count');
        return ctx.sync()
            .then(function()
            {
                output("Bindings count: " + bindings.count);
                return bindings.count;
            })
            .catch(function(error)
            {
                output("Error: " + error);
                if (error instanceof OfficeExtension.Error)
                {
                    output("Debug info: " + JSON.stringify(error.debugInfo));
                }
                return -1;
            });
    });
}

function getBindingsInfo()
{
    var outputString = "";

    Excel.run(function(ctx) 
    {
        var bindings = ctx.workbook.bindings;
        bindings.load('items');
        return ctx.sync().then(function()
        {
            outputString = "Count: " + bindings.items.length + ", "
            for (var i = 0; i < bindings.items.length; i++)
            {
                var binding = bindings.items[i];
                outputString = outputString + "{ Id: " + binding.id + ", Type: " + binding.type + " } ";
            }
            output(outputString);
        })
        .catch(function(error)
        {
            output("Error: " + error);
            if (error instanceof OfficeExtension.Error)
            {
                output("Debug info: " + JSON.stringify(error.debugInfo));
            }
            return -1;
        });
    });
}

function returnValueToAutomation(res)
{
    if (res instanceof Microsoft.Office.WebExtension.TableData)
    {
        var rows = res.rows;
        var headers = res.headers
        output("rows: " + rows + "||headers: " + headers);
    }
    else
    {
        output(res);
    }
}




// ========== Functions: Setting ==========
function setSettings()
{
    if (undefined === _Settings) 
    {
        output("Set settings failed");
        return;
    }

    output("");
    for (i = 0; i < names.length; i++)
    {
        _Settings.set(names[i], values[i]);
    }
    output("Set settings complete");
}

function getSettings()
{
    if (undefined === _Settings) 
    {
        output("Get settings failed");
        return;
    }

    result = "";
    for (i = 0; i < names.length; i++)
    {
        result = result + _Settings.get(names[i]) + "\n";
    }
    output("Get settings returned : " + result);
}

function saveSettings()
{
    _Settings.saveAsync(
        function (asyncResult)
        {
            if(asyncResult.status == Office.AsyncResultStatus.Failed)
            {
                output("Save settings failed with error = " + asyncResult.error.name + ":" + asyncResult.error.message);
            }
        else
        {
            output("Save settings complete");
        }
        }
    );
}

function removeSettings()
{
    if (undefined === _Settings) 
    {
        output("Remove settings failed");
        return;
    }

    result = "";
    for (i = 0; i < names.length; i++)
    {
        _Settings.remove(names[i]);
        result = result + values[i] + "\n";
    }
    output("Removed elements : " + result);
}

function refreshSettings()
{
    _Settings.refreshAsync(
        function (asyncResult)
        {
            if (asyncResult.status == Office.AsyncResultStatus.Failed)
            {
                output("Refresh settings failed with error = " + asyncResult.error.name + ":" + asyncResult.error.message);
            }
            else
            {
                output("Refresh settings complete");
            }
        }
   );
}

// ========== Functions: Misc ==========
function output(str)
{
    var outputArea = document.getElementById("output");
    outputArea.value = str;

    var logsArea = document.getElementById("logs");
    logsArea.value = getTimestamp() + " - " + str + "\n" + logsArea.value;
}

function getTimestamp()
{
    var dt = new Date(Date.now());
    var hours = dt.getHours();
    var minutes = "0" + dt.getMinutes();
    var seconds = "0" + dt.getSeconds();

    return hours + ':' + minutes.substr(-2) + ':' + seconds.substr(-2);
}