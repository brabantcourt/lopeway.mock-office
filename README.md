lopeway.mock-office
===================

An attempt to mock the [JavaScript API for Office](//msdn.microsoft.com/en-us/library/fp142185.aspx) in order to facilitate debugging using a regular browser and its tools. The implementation is currently limited to part of the implementation of the [TableBinding](http://msdn.microsoft.com/en-us/library/fp160977.aspx) object although much of the infrastructure to support this object will support the creation of the additional objects.

# Installation

If you don't already, your project will need to reference [jQuery](//jquery.com) and [KnockoutJS](//knockoutjs.com) as these are required by the spreadsheet 'emulator'.

    <script src="//ajax.aspnetcdn.com/ajax/jQuery/jquery-1.8.1.min.js"></script>
    <script src="//ajax.aspnetcdn.com/ajax/knockout/knockout-2.1.0.js"></script>

Replace the reference to Microsoft's JavaScript API for Office:

    <!--<script src="//.../office.js"></script>-->

With a reference to the mock version:

	<script src ="//.../lopeway.office.mock.1.2.js"></script>

# Configuration

To override the mock library's default settings you must precede the lopeway.office.mock file reference to a configuration JSON of the form:

	lopeway = {
        excel: {
            config: {
                debug: true,
                id: "lopeway-workbook",
                sheets: "sheets",
                tables: "tables",
                workbook: {
                    name: "Book1",
                    sheets: [],
                    tables: []
                }
            }
        }
    }

<table>
	<tr><td>debug</td><td>'true' to display tables/sheets as they're rendered; 'false' to hide the display</td></tr>
	<tr><td>id</td><td>ID of the DIV to which rendered content will be appended</td></tr>
	<tr><td>sheets</td><td>ID of the DIV to which rendered sheets will be appended</td></tr>
	<tr><td>tables</td><td>ID of the DIV to which rendered tables will be appended</td></tr>
	<tr><td>workbook</td><td>Definition of the Excel spreadsheet</td></tr>
	<tr><td>name</td><td>Unused</td></tr>
	<tr><td>sheets</td><td>Format { name: 'sheetname' }</td></tr>
	<tr><td>tables</td><td>Format { name: 'tablename', cols = [c1, c2...], rows = [[r1c1,r1c2...],[r2c1,r2c2...] }</td></tr>
</table>

# Unit Test(s)

[/tests/lopeway.office.mock.1.3.unittests.js](http://github.com/brabantcourt/lopeway.mock-office/blobk/master/tests/lopeway.office.mock.1.3.html)

# Example(s)

[/examples/default.html](http://github.com/brabantcourt/lopeway.mock-office/blobk/master/examples/default.html)