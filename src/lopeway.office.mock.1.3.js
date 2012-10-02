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

    var config, bookModel, viewModel, ViewModel, Cols, Rows, Sheet, Table, Workbook;

    // User overrides by predefining lopeway.excel.confg
    // Config minimally requires id, debug, sheets, tables
    // Setting defaults for each of these if not present
    config = (lopeway && lopeway.excel && lopeway.excel.config) || {};
    config.id || (config.id = "lopeway-workbook");
    config.debug || (config.debug = false);
    config.sheets || (config.sheets = "sheets");
    config.tables || (config.tables = "tables");
    config.workbook || (config.workbook = { name: "Book1" });

    // Add self into the page
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
        function Sheet(name) {
            this.name = name;
        }
        Sheet.prototype.get = function () {
            // Created for symmetry
            // Currently redundant since Sheet's properties aren't defined
            console.warn("[Sheet:get] not yet implemented!");
        };
        Sheet.prototype.set = function () {
            if (viewModel.sheet.index(this.name) !== -1) {
                // Sheet name exists===T in view model
                console.warn("[Sheet] overwriting existing sheet!");
            } else {
                // Sheet name exists!==T in view model
                viewModel.sheet.set(this.name);
            }
        };
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
            this.cols = new Cols(this.name);
            this.rows = new Rows(this.name);
        }
        Table.prototype.get = function () {
            // Created for symmetry
            // Currently redundant since Table's properties exclusively represented by its Cols/Rows
            console.warn("[Table:get] not yet implemented!");
        };
        Table.prototype.set = function () {
            if (viewModel.table.index(this.name) !== -1) {
                // Table name exists===T in view model
                console.warn("[Table] overwriting existing table!");
            } else {
                // Table name exists!==T in view model
                viewModel.table.set(this.name);
            }
        };
        Table.prototype.exists = function () {
            return viewModel.table.index(this.name) !== -1;
        };
        return Table;
    }());
    Workbook = (function () {
        function Workbook(workbook) {
            this.name = workbook.name;
            // if there are no workbook.sheets/tables, create empty sets
            workbook.sheets || (workbook.sheets = []);
            workbook.tables || (workbook.tables = []);
            // iterate over them
            workbook.sheets.forEach(function (sheet) {
                var s = new Sheet(sheet.name);
                s.set();
                // No additional properties to set
            });
            workbook.tables.forEach(function (table) {
                var t = new Table(table.name);
                t.set();
                if (table.cols) { t.cols.set(table.cols); }
                if (table.rows) { t.rows.set(table.rows); }
            });
        }
        return Workbook;
    }());

    // Create and open the Workbook
    // Because Workbook is a veneer over the underlying view model, can discard the workbook object once constructed
    var workbook = new Workbook(config.workbook);

    // Push it into the global namespace
    window.lopeway = window.lopeway || {};
    window.lopeway.excel = window.lopeway.excel || {};
    window.lopeway.excel.config || (window.lopeway.excel.config = config);
    window.lopeway.excel.mock= {
        Sheet: Sheet,
        Table: Table
    };

});
(function () {

    "use strict";

    var __hasProp, __extends, bindings = [], tablebindings = [], Binding, MatrixBinding, TableBinding, TextBinding;

    // CoffeeScript compiler
    __hasProp = Object.prototype.hasOwnProperty;
    __extends = function (child, parent) {
        var key;
        for (key in parent) {
            if (__hasProp.call(parent, key)) {
                child[key] = parent[key];
            }
        }
        function ctor() {
            this.constructor = child;
        }
        ctor.prototype = parent.prototype;
        child.prototype = new ctor;
        child.__super__ = parent.prototype;
        return child;
    };

    Binding = (function () {
        function Binding(name) {
            console.warn("[Binding] not yet implemented!");
            this.document = null;
            this.id = null;
            this.type = null;
        }
        Binding.prototype.addHandlerAsync = function () {

        };
        Binding.prototype.getDataAsync = function () {

        };
        Binding.prototype.removeHandlerAsync = function () {

        };
        Binding.prototype.setDataAsync = function () {

        };
        return Binding;
    })();
    MatrixBinding = (function (_super) {
        __extends(MatrixBinding, Binding);
        function MatrixBinding(name) {
            console.warn("[MatrixBinding] not yet implemented!");
            MatrixBinding.__super__.constructor.call(this, name);
        }
        return MatrixBinding;
    }());
    TableBinding = (function (_super) {
        __extends(TableBinding, Binding);
        function TableBinding(name) {
            TableBinding.__super__.constructor.call(this, name);
            this.table = new lopeway.excel.mock.Table(name);
            this.rowCount = 0;
        }
        TableBinding.prototype.addRowsAsync = function (rows) {
            var underlyingRows;
            underlyingRows = this.table.rows.get();
            rows.forEach(function (row) { underlyingRows.push(row); });
            this.table.rows.set(underlyingRows);
            this.rowCount = underlyingRows.length;
        };
        TableBinding.prototype.deleteAllDataValuesAsync = function () {
            this.table.rows.set([]);
            this.rowCount = 0;
        };
        return TableBinding;
    }());
    TextBinding = (function (_super) {
        __extends(TextBinding, Binding);
        function TextBinding(name) {
            console.warn("[TextBinding] not yet implemented!");
            TextBinding.__super__.constructor.call(this, name);
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