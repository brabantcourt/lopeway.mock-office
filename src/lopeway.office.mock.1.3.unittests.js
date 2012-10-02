/*globals module,asyncTest,test,expect,equal,deepEqual,ok,raises,start,stop*/
$(function () {

    "use strict";

    module("lopeway.excel.mock");

    var config = lopeway.excel.config, excel = lopeway.excel.mock;

    test("jQuery", function () {
        // Expect
        expect(3);
        // Arrange
        // Act
        // Assert
        equal($("#{0}".format(config.id)).length, 1);
        equal($("#{0} #{1}".format(config.id, config.sheets)).length, 1);
        equal($("#{0} #{1}".format(config.id, config.tables)).length, 1);
    })

    test("Sheet", function () {
        // Expect
        expect(5);
        // Arrange
        // Act
        var name = "SheetTest", sheet;
        // Assert

        equal($("#{0} #{1} > tbody > tr".format(config.id, config.sheets)).length, config.workbook.sheets.length);

        sheet = new excel.Sheet(name);
        ok(!sheet.exists());
        sheet.set();
        ok(sheet.exists());

        equal($("#{0} #{1} > tbody > tr".format(config.id, config.sheets)).length, config.workbook.sheets.length + 1);
        equal($("#{0} #{1} > tbody > tr td[data-bind=\"text:name\"]:contains(\"{2}\")".format(config.id, config.sheets, name)).length, 1);

    });
    test("Table", function () {
        // Expect
        expect(13);
        // Arrange
        // Act
        var name = "Test1!Test1", table, $tbody;
        // Assert

        equal($("#{0} #{1} > tbody > tr".format(config.id, config.tables)).length, config.workbook.tables.length);

        table = new excel.Table(name);
        ok(!table.exists());
        table.set();
        ok(table.exists());

        equal($("#{0} #{1} > tbody > tr".format(config.id, config.tables)).length, config.workbook.tables.length + 1);
        equal($("#{0} #{1} > tbody > tr td[data-bind=\"text:name\"]:contains(\"{2}\")".format(config.id, config.tables, encodeURIComponent(name.replace("!", "_")))).length, 1);

        table.rows.set(null);
        equal(table.rows.get(), null);
        table.rows.set([]);
        equal(table.rows.get().length, 0);
        table.rows.set([[]]);
        equal(table.rows.get().length, 1);

        // Assign '$tbody' to the HTML/DOM node containing 'table'
        $tbody = $("#{0} #{1} > tbody > tr td[data-bind=\"text:name\"]:contains(\"{2}\")".format(config.id, config.tables, encodeURIComponent(name.replace("!", "_")))).siblings().find("table tbody");
        equal($tbody.find("tr").length, 1);
        table.rows.set([[],[]])
        equal($tbody.find("tr").length, 2);
        table.rows.set([[1]]);
        equal($tbody.find("tr").length, 1);
        table.rows.set([[1,2]]);
        equal($tbody.find("tr").length, 1);
        table.rows.set([[1, 2], [1, 2]]);
        equal($tbody.find("tr").length, 2);

    });

    module("lopeway.office.mock");

    var office = lopeway.office.mock;

    test("Binding", function () {
        // Expect
        expect(0);
        // Arrange
        // Act
        var name="Sheet!Data1", binding = new office.Binding(name);
        // Assert

    })
    test("TableBinding", function () {
        // Expect
        expect(1);
        // Arrange
        // Act
        var name = "Sheet1!Data1", tablebinding = new office.TableBinding(name);
        // Assert
        equal(tablebinding.rowCount, 0);
        /*
        tablebinding.addRowsAsync([]);
        equal(tablebinding.rowCount, 0);
        tablebinding.addRowsAsync([[]]);
        equal(tablebinding.rowCount, 1);
        tablebinding.addRowsAsync([[],[]]);
        equal(tablebinding.rowCount, 2);
        */
    });

});
