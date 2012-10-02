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

    var bookModel, viewModel, ViewModel, Cols, Rows, Sheet, Table, Workbook;

    // Move this out of here to permit user to more easily set parameters
    var config = {
        debug: true,
        id: "lopeway-workbook",
        sheets: "sheets",
        tables: "tables",
        workbook: {
            name: "Workbook1.xlsx",
            sheets: [{
                name: "Sheet1"
            }, {
                name: "Sheet2"
            }],
            tables: [{
                name: "Sheet1!Data",
                cols: ["ID", "Name", "Lat", "Lng"],
                rows: [[1, "Company #1", 0.00, 0.00], [2, "Company #2", 0.00, 0.00], [3, "Company #3", 0.00, 0.00], [4, "Company #4", 0.00, 0.00], [5, "Company #5", 0.00, 0.00]]
            }, {
                name: "Sheet2!Config",
                cols: ["ID"],
                rows: [["sorrentos-coffee"], ["peets-coffee-redmond"], ["victors-coffee-company-redmond"]]
            },
            {
                name: "Data!Data",
                cols:["ID","Name","Rating","#Reviews","Lat","Lng"],
                rows:[[1,"Company #1",null,0,0.00,0.00]]
            }]
        }
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
                            <thead data-bind="foreach:cols">
                                <tr>
                                    <!-- ko foreach: $data -->
                                    <td data-bind="text:$data"></td>
                                    <!-- /ko -->                                        
                                </tr>                             
                            </thead>
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

    // Intentionally mirrors (as it does functionally) the approach taken with the KnockoutJS viewModel below
    bookModel = {
        sheets: function () {
            var row, rows, body;
            rows = function () {
                return $("<table></table>");
            };
            row = $("<tr></tr>").append($("<td></td>").attr("data-bind", "text:name"), $("<td></td>").append(rows()));
            body = $("<tbody></tbody>").attr("data-bind", "foreach:sheets").append(row);
            return $("<table></table>").attr("id", config.sheets).append(body);
        },
        tables: function () {
            var row, rows, body;
            rows = function () {
                var row, body, head;
                // Uses append(x1,x2,x3) to add three items (two of which are comment bindings) to the header row
                row = $("<tr></tr>").append($("<!-- ko foreach: cols -->"), $("<th></th>").attr("data-bind", "text:$data"), $("<!-- /ko -->"));
                head = $("<thead></thead>").append(row);
                // Uses append(x1,x2,x3) to add three items (two of which are comment bindings) to the body row
                row = $("<tr></tr>").append($("<!-- ko foreach: $data -->"), $("<td></td>").attr("data-bind", "text:$data"), $("<!-- /ko -->"));
                body = $("<tbody></tbody>").attr("data-bind", "foreach:rows").append(row);
                return $("<table></table>").append(head, body);
            };
            // Uses append(x1,x2) to add two TDs the first binds to the 'name' property, the second contains the table representing the rows
            row = $("<tr></tr>").append($("<td></td>").attr("data-bind", "text:name"), $("<td></td>").append(rows()));
            body = $("<tbody></tbody>").attr("data-bind", "foreach:tables").append(row);
            return $("<table></table>").attr("id", config.tables).append(body);
        }
    };
    $("#{0}".format(config.id)).append(bookModel.tables(), bookModel.sheets());

    ViewModel = (function () {
        var Sheet, Table;
        Sheet = (function () {
            function Sheet(sheets) {
                this.sheets = sheets;
            }
            Sheet.prototype.index = function (name) {
                return this.sheets().indexOfProperty("name", name);
            };
            Sheet.prototype.set = function (name) {
                this.sheets.push({
                    name: name
                });
            };
            Sheet.prototype.get = function (name) {
                return this.sheets()[this.sheets().indexOfProperty("name", name)];
            };
            return Sheet;
        }());
        Table = (function () {
            function Table(tables) {
                this.tables = tables;
            }
            Table.prototype.index = function (name) {
                return this.tables().indexOfProperty("name", name);
            };
            Table.prototype.set = function (name) {
                this.tables.push({
                    name: name,
                    cols: ko.observableArray([]),
                    rows: ko.observableArray([])
                });
            };
            Table.prototype.get = function (name) {
                return this.tables()[this.tables().indexOfProperty("name", name)];
            };
            return Table;
        }());
        function ViewModel() {
            this.sheets = ko.observableArray([]);
            this.tables = ko.observableArray([]);
            this.sheet = new Sheet(this.sheets);
            this.table = new Table(this.tables);
        }
        return ViewModel;
    }());
    viewModel = new ViewModel();
    ko.applyBindings(viewModel, $("#{0}".format(config.id)).get(0));

    Sheet = (function () {
        function Sheet(o) {
            this.name = o.name;
            viewModel.sheet.set(this.name);
        }
        Sheet.prototype.exists = function () {
            return viewModel.sheet.index(this.name) !== -1;
        };
        return Sheet;
    }());
    Cols = (function () {
        function Cols(name) {
            this.name = name;
        }
        Cols.prototype.get = function () {
            viewModel.table.get(this.name).cols();
        };
        Cols.prototype.set = function (cols) {
            viewModel.table.get(this.name).cols(cols);
        };
        return Cols;
    }());
    Rows = (function () {
        function Rows(name) {
            this.name = name;
        }
        Rows.prototype.get = function () {
            return viewModel.table.get(this.name).rows();
        };
        Rows.prototype.set = function (rows) {
            viewModel.table.get(this.name).rows(rows);
        };
        return Rows;
    }());
    Table = (function () {
        function Table(name) {
            if (name.indexOf("!") === -1) {
                console.warn("[Table] JavaScript API for Office requires that table names include the sheet reference, e.g. \"Sheet1!Table1\". This table's name ({0}) does not.".format(name));
            }
            // Ensure that "Sheet1!Data" --> Sheet1_Data so that it represents a valid HTML ID
            this.name = encodeURIComponent(name.replace("!", "_"));

            // The viewModel.tables acts as a persistence layer for Table(s)
            // If this table is not in the view model, create it
            if (viewModel.table.index(this.name) === -1) {
                viewModel.table.set(this.name);
            }

            // Interaction with the table occurs through its cols/rows
            // Create references for them
            // Nothing is actually done beyond persisting their table's name
            this.cols = new Cols(this.name);
            this.rows = new Rows(this.name);

        }
        return Table;
    }());
    Workbook = (function () {
        function Workbook(workbook) {
            workbook.sheets.forEach(function (sheet) {
                console.warn("[Workbook] Constructor still uses old-form 'Sheets'");
                var s = new Sheet(sheet);
            });
            workbook.tables.forEach(function (table) {
                var t = new Table(table.name);
                if (table.cols) { t.cols.set(table.cols); }
                if (table.rows) { t.rows.set(table.rows); }
            });
        }
        return Workbook;
    }());

    // This may be better placed in the "Office" section, but...
    // Model the Workbook
    var workbook = new Workbook(config.workbook);

    // Push it into the global namespace
    window.lopeway = window.lopeway || {};
    window.lopeway.excel = window.lopeway.excel || {};
    window.lopeway.excel.mock = {
        Sheet: Sheet,
        Table: Table
    };

    // Expose 'private' properties for unit testing
    window.lopeway.excel.unittests = {
        config: config
    };

});
(function () {

    "use strict";

    var bindings = [], tablebindings = [], Binding, MatrixBinding, TableBinding, TextBinding;

    Binding = (function () {
        function Binding(name) {
            console.warn("[Binding] not yet implemented!");
        }
        return Binding;
    }());
    MatrixBinding = (function () {
        function MatrixBinding(name) {
            console.warn("[MatrixBinding] not yet implemented!");
        }
        return MatrixBinding;
    }());
    TableBinding = (function () {
        function TableBinding(name) {
            this.table = new lopeway.excel.mock.Table(name);
        }
        TableBinding.prototype.addRowsAsync = function (rows) {
            var underlyingRows;
            underlyingRows = this.table.rows.get();
            rows.forEach(function (row) { underlyingRows.push(row); });
            this.table.rows.set(underlyingRows);
        };
        TableBinding.prototype.deleteAllDataValuesAsync = function () {
            this.table.rows.set([]);
        };
        return TableBinding;
    }());
    TextBinding = (function () {
        function TextBinding() {
            console.warn("[TextBinding] not yet implemented!");
        }
        return TextBinding;
    }());

    window.Office = {};
    window.Office.AsyncResultStatus = {
        Failed: "failed",
        Succeeded: "succeeded"
    };
    window.Office.BindingType = {
        Matrix: "matrix",
        Table: "table",
        Text: "text"
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
            var asyncResult = {
                // Assume the worst and be proven otherwise
                status: window.Office.AsyncResultStatus.Failed
            };
            console.debug("[addFromNamedItemAsync]");

            switch (type) {
                case window.Office.BindingType.Matrix:
                    console.warn("[addFromNamedItemAsync] type not yet implemented: {0}".format(type));
                    break;
                case window.Office.BindingType.Table:
                    console.debug("[addFromNamedItemAsync] TableBinding");
                    var table;
                    try {
                        table = new TableBinding(name);
                    } catch (notexist) {
                        console.warn("[addFromNamedItemAsync] caught something!");
                    }
                    asyncResult = {
                        status: window.Office.AsyncResultStatus.Succeeded,
                        value: table
                    };
                    break;
                case window.Office.BindingType.Text:
                    console.warn("[addFromNamedItemAsync] type not yet implemented: {0}".format(type));
                    break;
                default:
                    console.warn("[addFromNamedItemAsync] type undefined: {0}".format(type));
            }
            // Invoke the callback with asyncResult
            fn(asyncResult);
        }
    };

    // Simulate the slow loading of the JavaScript file then involve Office.initialize() if any
    setTimeout(function () {
        window.Office.initialize && window.Office.initialize();
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