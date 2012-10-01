/*globals module,asyncTest,test,expect,equal,deepEqual,ok,raises,start,stop*/
$(function () {

    "use strict";

    module("lopeway.excel.mock");

    var config = lopeway.excel.unittests.config,excel = lopeway.excel.mock;

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
        expect(1);
        // Arrange
        // Act
        var name = "Test1!Test1", sheet = new excel.Sheet(name);
        // Assert
        ok(sheet.exists());
    });
    test("Table", function () {
        // Expect
        expect(12);
        // Arrange
        // Act
        var name = "Test1!Test1", table;
        // Assert
        equal($("#{0} #{1} tbody tr".format(config.id, config.tables)).length, 0);
        table = new excel.Table(name);
        equal($("#{0} #{1} tbody tr".format(config.id, config.tables)).length, 1);
        equal($("#{0} #{1} tbody tr td[data-bind=\"text:name\"]".format(config.id, config.tables)).length, 1);
        equal($("#{0} #{1} tbody tr td[data-bind=\"text:name\"]".format(config.id, config.tables)).html(), encodeURIComponent(name.replace("!", "_")));
        ok(table.exists());
        table.setRows(null);
        equal(table.getRows(), null);
        table.setRows([]);
        equal(table.getRows().length, 0);
        table.setRows([[]]);
        equal(table.getRows().length, 1);
        equal($("#{0} #{1} tbody table tbody tr".format(config.id, config.tables)).length, 1);
        table.setRows([[],[]])
        equal($("#{0} #{1} tbody table tbody tr".format(config.id, config.tables)).length, 2);
        table.setRows([[1]]);
        equal($("#{0} #{1} tbody table tbody tr td".format(config.id, config.tables)).length, 1);
        table.setRows([[1,2]]);
        equal($("#{0} #{1} tbody table tbody tr td".format(config.id, config.tables)).length, 2);
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
