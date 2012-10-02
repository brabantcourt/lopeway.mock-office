/*globals window*/
(function () {

    "use strict";

    var lopeway = {
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
    };

    window.lopeway = lopeway;

}());