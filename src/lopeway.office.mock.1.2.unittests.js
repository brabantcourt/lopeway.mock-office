/*globals module,asyncTest,test,expect,equal,deepEqual,ok,raises,start,stop*/
$(function () {

    "use strict";

    module("lopeway.excel.mock");

    var config = lopeway.excel.unittests.config, excel = lopeway.excel.mock;

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
        expect(4);
        // Arrange
        // Act
        var name = "SheetTest", sheet;
        // Assert
        equal($("#{0} #{1} > tbody > tr".format(config.id, config.sheets)).length, config.workbook.sheets.length);
        sheet = new excel.Sheet({ name: "{0}".format(name) });
        equal($("#{0} #{1} > tbody > tr".format(config.id, config.sheets)).length, config.workbook.sheets.length + 1);
        equal($("#{0} #{1} > tbody > tr td[data-bind=\"text:name\"]:contains(\"{2}\")".format(config.id, config.sheets, name)).length, 1);
        ok(sheet.exists());
    });
    test("Table", function () {
        // Expect
        expect(12);
        // Arrange
        // Act
        var name = "Test1!Test1", table, $tbody;
        // Assert
        equal($("#{0} #{1} > tbody > tr".format(config.id, config.tables)).length, config.workbook.tables.length);

        table = new excel.Table(name);
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
        var binding = new office.Binding("");
        // Assert
    })
    test("TableBinding", function () {
        // Expect
        expect(0);
        // Arrange
        // Act
        var name = "Test1!Test1", tablebinding = new office.TableBinding("");
        // Assert

    });

});
