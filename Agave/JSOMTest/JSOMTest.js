// Originated from
// src\osftest\files\MOEs\JSOMSettings.js, \\daddev\Office\16.0\19104.12000\src\osftest\files\MOEs
// JSOMEvents_ExcelIncontent.js

// Reference docs
// https://docs.microsoft.com/en-us/office/dev/add-ins/excel/excel-add-ins-advanced-concepts

var result;

// ========== Variables: Binding ==========
var _bindingType;
var _bindingEventHandlers = [];


// ========== Variables: Setting ==========
var _Bindings;
var _Settings;
var names = ['StringSetting', 'StatusSetting', 'IntegerSetting', 'ArrayOfIntegersSetting'];
var values = [];
values[0] = 'This is a string value'; //string value
values[1] = true; //boolean value
values[2] = 45; //number
values[3] = [5, 8, 10, 100]; // array


// ========== Init ==========
Office.onReady(function(info)
{
    OfficeExtension.config.extendedErrorLogging = true;

    _Settings = Office.context.document.settings;
    _Bindings = Office.context.document.bindings;
    if (undefined === _Settings) 
    {
        return;
    }

    //registerSettingsChangedEventHandler();
    //registerBindingEventHandlersOnLoad();
    displayCurrentTime();
});

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

    Excel.run(function(ctx) 
    {
        ctx.workbook.bindings.addFromSelection(_bindingType, bindingId);
        return ctx.sync().then(function()
        {
            output("Added new binding, id: " + bindingId + ", type: " + _bindingType);
        });
    }).catch(function(error) { logError(error); });
}

function removeBinding()
{
    var bindingId = getBindingId();
    if (!bindingId || bindingId.trim().length == 0)
    {
        output("Error: Need to specify binding id to remove");
        return;
    }

    Excel.run(function(ctx) 
    {
        var binding = ctx.workbook.bindings.getItem(bindingId);
        return ctx.sync().then(function()
        {
            binding.delete();
            return ctx.sync().then(function()
            {
                output("Removed binding, id: " + bindingId);
            });
        });
    }).catch(function(error) { logError(error); });
}

function getBindingsCount()
{
    Excel.run(function(ctx) 
    {
        var bindings = ctx.workbook.bindings;
        bindings.load('count');
        return ctx.sync().then(function()
        {
            output("Bindings count: " + bindings.count);
            return bindings.count;
        });
    }).catch(function(error)
    {
        logError(error);
        return -1;
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
            for (var i = 0; i < bindings.items.length; i++)
            {
                var binding = bindings.items[i];
                outputString = outputString + "{ Id: " + binding.id + ", Type: " + binding.type + " }" + "\r\n";
            }

            output(outputString);
        });
    }).catch(function(error) { logError(error); });
}

function getBindingRange()
{
    var bindingId = getBindingId();
    if (!bindingId || bindingId.trim().length == 0)
    {
        output("Error: Need to specify binding id");
        return;
    }

    Excel.run(function(ctx) 
    {
        var binding = ctx.workbook.bindings.getItem(bindingId);
        binding.load("type, text");
        return ctx.sync().then(function()
        {
            if (binding.type == "Text")
            {
                var text = binding.getText();
                ctx.sync().then(function()
                {
                    output("id: " + bindingId + ", text: " + text.value);
                });
            }
            else if (binding.type == "Range")
            {
                var bindingRange = binding.getRange();
                bindingRange.load(['address', 'cellCount']);
                return ctx.sync().then(function()
                {
                    output("id: " + bindingId + ", address: " + bindingRange.address + ", cellCount: " + bindingRange.cellCount);
                    var sheet = ctx.workbook.worksheets.getActiveWorksheet();
                    var range = sheet.getRange(bindingRange.address);
                    range.select();
                    return ctx.sync();
                });
            }
            else
            {
                var table = binding.getTable();
                var tableRange = table.getRange();
                tableRange.load(['address', 'cellCount']);
                return ctx.sync().then(function()
                {
                    output("id: " + bindingId + ", address: " + tableRange.address + ", cellCount: " + tableRange.cellCount);
                    tableRange.select();
                    return ctx.sync();
                });
            }
        });
    }).catch(function(error) { logError(error); });
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
        result = result + "id: " + names[i] + ", value: " + _Settings.get(names[i]) + "\n";
    }
    output("Get settings returned : \r\n" + result);
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

// ========== Functions: Event Handlers ==========
function registerSettingsChangedEventHandler()
{
    Excel.run(function(ctx) 
    {
        var settings = ctx.workbook.settings; 
        settings.onSettingsChanged.add(handleSettingsChanged);
        return ctx
            .sync()
            .then(function()
            {
                output("Settings changed event handler registered");
            });
    }).catch(function(error) { logError(error) });
}

function handleSettingsChanged(eventArgs)
{
    output("SettingsChanged Event - Calling refreshSettings()");
    refreshSettings();
}

// NOT USED FOR NOW
function registerBindingEventHandlersOnLoad()
{
    // TODO: Is this right?
    // TODO: Remove existing event handlers
    Excel.run(function(ctx) 
    {
        var bindings = ctx.workbook.bindings;
        bindings.load('items');
        return ctx.sync().then(function()
        {
            // Register event handlers
            for (var i = 0; i < bindings.items.length; i++)
            {
                var newEventHandler = bindings.items[i].onDataChanged.add(handleBindingDataChanged);
                _bindingEventHandlers.push(newEventHandler);
            }

            output("Registering binding event handlers... count: " + bindings.items.length);
            return ctx.sync().then(function()
            {
                output("Binding DataChanged handler registered");
            });
        });
    }).catch(function(error) { logError(error); });
}

function registerBindingEventHandler()
{
    Excel.run(function(ctx) 
    {
        var bindings = ctx.workbook.bindings;
        bindings.load('items');
        return ctx.sync().then(function()
        {
            var bindingId = getBindingId();
            for (var i = 0; i < bindings.items.length; i++)
            {
                var binding = bindings.items[i];
                if (binding.id == bindingId)
                {
                    var newEventHandler = binding.onDataChanged.add(handleBindingDataChanged);
                    return ctx.sync().then(function()
                    {
                        output("Binding DataChanged handler registered, id: " + bindingId);
                    });
                }    
            }
        });
    }).catch(function(error) { logError(error); });
}

function handleBindingDataChanged(eventArgs)
{
    var binding = eventArgs.binding;
    output("DataChange Event - BindingId: " + binding.id + ", BindingType: " + binding.type);
}

// ========== Functions: Misc ==========
function output(str)
{
    var outputArea = document.getElementById("output");
    outputArea.value = str;

    var logsArea = document.getElementById("logs");
    logsArea.value = getTimestamp() + " - " + str + "\n" + logsArea.value;
}

function logError(error)
{
    output("Error: " + error);
    if (error instanceof OfficeExtension.Error)
    {
        output("Debug info: " + JSON.stringify(error.debugInfo));
    }
}

function getTimestamp()
{
    var dt = new Date(Date.now());
    var hours = dt.getHours();
    var minutes = "0" + dt.getMinutes();
    var seconds = "0" + dt.getSeconds();

    return hours + ':' + minutes.substr(-2) + ':' + seconds.substr(-2);
}

function displayCurrentTime()
{
    var dt = getTimestamp();
    document.getElementById('currentDatetimeDiv').innerHTML = dt;
    var nextDt = setTimeout(displayCurrentTime, 1000);
}

function setBorderAroundSelectedRange()
{
    Excel.run(function (ctx)
    {
        var range = ctx.workbook.getSelectedRange();
        range.format.borders.getItem('EdgeBottom').style = 'Continuous';
        range.format.borders.getItem('EdgeBottom').weight = 'Thick';
        range.format.borders.getItem('EdgeLeft').style = 'Continuous';
        range.format.borders.getItem('EdgeLeft').weight = 'Thick';
        range.format.borders.getItem('EdgeRight').style = 'Continuous';
        range.format.borders.getItem('EdgeRight').weight = 'Thick';
        range.format.borders.getItem('EdgeTop').style = 'Continuous';
        range.format.borders.getItem('EdgeTop').weight = 'Thick';
        return ctx.sync(); 
    }).catch(function(error) { logError(error); });
}