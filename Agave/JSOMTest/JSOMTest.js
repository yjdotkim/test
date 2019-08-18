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
function getSettingKey()
{
    return document.getElementById("settingKey").value;
}

function setSettings()
{
    Excel.run(function(ctx) 
    {
        var settings = ctx.workbook.settings;
        for (i = 0; i < names.length; i++)
        {
            settings.add(names[i], values[i]);
        }
        return ctx.sync().then(function()
        {
            output("Set settings with sample key/values");
        });
    }).catch(function(error) { logError(error) });
}

function getSettings()
{
    Excel.run(function(ctx) 
    {
        var settings = ctx.workbook.settings;
        settings.load('items');
        return ctx.sync().then(function()
        {
            var result = "";
            var settingsCount = settings.items.length;
            output("Settings count: " + settingsCount);
            if (settingsCount > 0)
            {
                for (i = 0; i < settingsCount; i++)
                {
                    var setting = settings.items[i];
                    result = result + "key: " + setting.key + ", value: " + setting.value + "\n";
                }
            }

            output(result);
        });
    }).catch(function(error) { logError(error) });
}

function removeSetting()
{
    var key = getSettingKey();
    if (!key || key.trim().length == 0)
    {
        output("Error: Need to specify setting key to remove");
        return;
    }

    Excel.run(function(ctx)
    {
        var setting = ctx.workbook.settings.getItem(key);
        return ctx.sync().then(function()
        {
            setting.delete();
            return ctx.sync().then(function()
            {
                output("Removed setting, key: " + key);
            });
        });
    }).catch(function(error) { logError(error) });
}

// ========== Functions: Event Handlers ==========
function registerSettingsChangedEventHandler()
{
    Excel.run(function(ctx) 
    {
        var settings = ctx.workbook.settings; 
        settings.onSettingsChanged.add(handleSettingsChanged);
        return ctx.sync().then(function()
        {
            output("Settings changed event handler registered");
        });
    }).catch(function(error) { logError(error) });
}

function handleSettingsChanged(eventArgs)
{
    output("SettingsChanged Event fired");
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
    output("DataChange Event fired - BindingId: " + binding.id + ", BindingType: " + binding.type);
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