/*globals */
(function () {

    "use strict";

    String.prototype.format = function () {
        // fearphage's suggestion
        // http://stackoverflow.com/questions/610406/javascript-equivalent-to-printf-string-format/4673436#4673436
        var args = arguments;
        return this.replace(/\{(\d+)\}/g, function (match, number) {
            return typeof args[number] !== 'undefined' ? args[number] : match;
        });
    };

    Array.prototype.indexOfProperty = function (prop, value) {
        var arrayitem, _i, _len;
        for (_i = 0, _len = this.length; _i < _len; _i++) {
            arrayitem = this[_i];
            if (arrayitem[prop] === value) {
                return _i;
            }
        }
        return -1;
    };

}());
$(function () {


});
$(function () {

    "use strict";

    var config, div, mockoffice, viewModel;

    config = {
        id: "lopeway-workbook",
        sheets: "sheets",
        tables:"tables"
    };

    div = "<div></div>";

    // Add self into the page
    $("body").append($(div).attr("id", config.id));
    $("#{0}".format(config.id)).append($(div).attr("id", config.sheets).attr("data-bind", "foreach:sheets"));

    $("#{0}".format(config.id)).append($(div).attr("id", config.tables).attr("data-bind", "foreach:tables"));

    $("#{0} #{1}".format(config.id, config.tables)).append($(div).attr("class","table").attr("data-bind", "foreach:$data"));
    $("#{0} #{1} .table".format(config.id, config.tables)).append($(div).attr("data-bind", "text:name"));
    $("#{0} #{1} .table".format(config.id, config.tables)).append($(div).attr("data-bind", "text:$data"));

    viewModel = {
        sheets: ko.observableArray([]),
        tables: ko.observableArray([])
    };
    ko.applyBindings(viewModel, $("#{0}".format(config.id)).get(0));

    mockoffice = {
        addSheet: function (name) {
            // Make the name safe and replace any "!" references, so "Sheet1!Data"-->"Sheet1_Data"
            name = encodeURIComponent(name.replace("!", "_"));
            // [120930:Daz] Don't need to add this to the HTML since it can be effected through the view model
            //$("#{0} #{1}".format(config.id, config.sheets)).append($(div).attr("id", "sheet-{0}".format(name)).attr("data-bind","foreach:{0}".format(name)));
            viewModel.sheets.push({ name: name, data: ko.observableArray([]) });
        },
        addTable: function (name) {
            // Make the name safe and replace any "!" references, so "Sheet1!Data"-->"Sheet1_Data"
            name = encodeURIComponent(name.replace("!", "_"));
            // [120930:Daz] Don't need to add this to the HTML since it can be effected through the view model
            //$("#{0} #{1}".format(config.id, config.tables)).append($(div).attr("id", "table-{0}".format(name)).attr("data-bind","foreach:{0}".format(name)));
            viewModel.tables.push({ name: name, data: ko.observableArray([]) });
        },
        isSheet: function (name) {
            name = encodeURIComponent(name.replace("!", "_"));
            return viewModel.sheets().indexOfProperty("name", name);
        },
        isTable: function (name) {
            name = encodeURIComponent(name.replace("!", "_"));
            return viewModel.tables().indexOfProperty("name", name);
        }
    };

    // [120930:Daz] Make some tests here until I work out how to structure this more effectively!
    ["Sheet1!Data", "Sheet2!Config", "Sheet3!Errors"].forEach(function (name) {
        mockoffice.addTable(name);
        console.debug("isTable({0})={1}".format(name, mockoffice.isTable(name)));
    });


    // Push it into the global namespace
    window.lopeway = window.lopeway || {};
    window.lopeway.mockoffice = mockoffice;

});
(function () {

    window.Office = {};
    window.Office.AsyncResultStatus = {
        Failed: "failed",
        Succeeded: "succeeded"
    };
    window.Office.BindingType= {
        Table: "table"
    };
    window.Office.context= {
        document: {
            binding:{
                addRowsAsync: function (rows) {
                    console.log("[addRowsAsync]");
                }
            },
            bindings: {
                addFromNamedItemAsync: function (name, type, o, fn) {
                    console.debug("[addFromNamedItemAsync]");
                    // Either simulate successful binding by creating table or don't????
                    console.warn("Making assumption that the Excel table exists... perhaps add a choice?");
                    if (type === "table" && !lopeway.mockoffice.isTable[name]) { lopeway.mockoffice.addTable(name); }
                    // Don't forget to call the async function (w/ success)
                    console.warn("Making assumption about the return type of the async function");
                    fn({
                        status: window.Office.AsyncResultStatus.Succeeded,
                        value: ""
                    });
                }
            }
        }
    };

    // Simulate the slow loading of the JavaScript file then involve Office.initialize() if any
    setTimeout(function () {
        if (window.Office.initialize) {
            window.Office.initialize();
        }
    }, 1000);

}());