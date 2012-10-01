/*globals $,console,ko,lopeway,setTimeout,window*/
/*jslint nomen: true*/
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

    "use strict";

    var config, viewModel, Sheet, Table;

    config = {
        debug: false,
        id: "lopeway-workbook",
        sheets: "sheets",
        tables: "tables"
    };

    // Add self into the page
    /*
    <div id="lopeway-workbook">
    <table id="sheets">
    <tbody data-bind="foreach:sheets">
    </tbody>
    </table>
    <table id="tables">
    <tbody data-bind="foreach:tables" >
    <tr >
    <td data-bind="text:name"></td>
    <td>
    <table>
    <tbody data-bind="foreach:rows">
    <tr>
    <!-- ko foreach: $data -->
    <td data-bind="text:$data"></td>
    <!-- /ko -->                                        
    </tr>    
    </tbody>
    </table>
    </td>
    </tr>
    </tbody>
    </table>
    </div>
    */

    $("body").append($("<div></div>").attr("id", config.id).css("display", config.debug ? "block" : "none"));

    $("#{0}".format(config.id)).append($("<table></table>").attr("id", config.sheets).attr("data-bind", "foreach:sheets"));
    // see below... need to reproduce for 'sheets'

    $("#{0}".format(config.id)).append($("<table></table").attr("id", config.tables));
    $("#{0} #{1}".format(config.id, config.tables)).append($("<tbody></tbody").attr("data-bind", "foreach:tables"));
    $("#{0} #{1} tbody".format(config.id, config.tables)).append($("<tr></tr>"));
    $("#{0} #{1} tbody tr".format(config.id, config.tables)).append($("<td></td>").attr("data-bind", "text:name"));
    $("#{0} #{1} tbody tr".format(config.id, config.tables)).append($("<td><table></table></td>"));
    $("#{0} #{1} tbody tr td table".format(config.id, config.tables)).append($("<tbody></tbody>").attr("data-bind", "foreach:rows"));
    $("#{0} #{1} tbody tr td table tbody".format(config.id, config.tables)).append($("<tr><!-- ko foreach: $data --><td data-bind=\"text:$data\"></td><!-- /ko --></tr>"));

    viewModel = {
        sheets: ko.observableArray([]),
        tables: ko.observableArray([])
    };
    viewModel.getSheetIndex = function (name) { return viewModel.sheets().indexOfProperty("name", name); };
    viewModel.getSheet = function (name) { return viewModel.sheets()[viewModel.sheets().indexOfProperty("name", name)]; };
    viewModel.getTableIndex = function (name) { return viewModel.tables().indexOfProperty("name", name); };
    viewModel.getTable = function (name) { return viewModel.tables()[viewModel.tables().indexOfProperty("name", name)]; };
    ko.applyBindings(viewModel, $("#{0}".format(config.id)).get(0));

    Sheet = (function () {
        function Sheet(name) {
            this.name = encodeURIComponent(name.replace("!", "_"));
            viewModel.sheets.push({
                name: this.name
                /* data? */
            });
        }
        Sheet.prototype.exists = function () {
            return viewModel.getSheetIndex(this.name) !== -1;
        };
        return Sheet;
    } ());
    Table = (function () {
        function Table(name) {
            this.name = encodeURIComponent(name.replace("!", "_"));
            viewModel.tables.push({
                name: this.name,
                cols: ko.observableArray([]),
                rows: ko.observableArray([])
            });
        }
        Table.prototype.exists = function () {
            return viewModel.getTableIndex(this.name) !== -1;
        };
        Table.prototype.setRows = function (rows) {
            viewModel.getTable(this.name).rows(rows);
        };
        Table.prototype.getRows = function () {
            return viewModel.getTable(this.name).rows();
        };
        return Table;
    } ());

    // Push it into the global namespace
    window.lopeway = window.lopeway || {};
    window.lopeway.excel = window.lopeway.excel || {};
    window.lopeway.excel.mock = {
        Sheet: Sheet,
        Table: Table
    };

    window.lopeway.excel.unittests = {
        config: config
    };

});
(function () {

	"use strict";

	var bindings = [], tablebindings = [], Binding, MatrixBinding, TableBinding;

    Binding = (function () {
        function Binding(name) {
        }
        return Binding;
    }());
    MatrixBinding=(function(){
        function MatrixBinding(name){
            
        }
        return MatrixBinding;
    }());
    TableBinding = (function () {
        function TableBinding(name) {
            this.table = new lopeway.excel.mock.Table(name);
        }
        TableBinding.prototype.addRowsAsync = function (rows) {
            var underlyingRows;
            underlyingRows = this.table.getRows();
            rows.forEach(function (row) { underlyingRows.push(row); });
            this.table.setRows(underlyingRows);
        };
        TableBinding.prototype.deleteAllDataValuesAsync = function () {
            this.table.setRows([]);
        };
        return TableBinding;
    }());

    window.Office = {};
    window.Office.AsyncResultStatus = {
        Failed: "failed",
        Succeeded: "succeeded"
    };
    window.Office.BindingType = {
        Table: "table"
    };
    window.Office.context = {};
    window.Office.context.document = {};
    window.Office.context.document.binding = {
        addRowsAsync: function (rows) {
            console.warn("[addRowsAsync] not yet implemented");
        },
        deleteAllDataValuesAsync: function () {
            console.warn("[deleteAllDataValuesAsync] not yet implemented");
        }
    };
    window.Office.context.document.bindings = {
        addFromNamedItemAsync: function (name, type, o, fn) {
            console.debug("[addFromNamedItemAsync]");
            // Either simulate successful binding by creating table or don't????
            console.warn("Making assumption that the Excel table exists... perhaps add a choice?");

            console.warn("Making assumption about the return type of the async function");
            fn({
                status: window.Office.AsyncResultStatus.Succeeded,
                value: new TableBinding(name)
            });
        }
    };

    // Simulate the slow loading of the JavaScript file then involve Office.initialize() if any
    setTimeout(function () {
        if (window.Office.initialize) {
            window.Office.initialize();
        }
    }, 1000);

        // Push it into the global namespace
    window.lopeway = window.lopeway || {};
    window.lopeway.office = window.lopeway.office || {};
    window.lopeway.office.mock = {
        Binding: Binding,
        MatrixBinding: MatrixBinding,
        TableBinding: TableBinding
    };

}());