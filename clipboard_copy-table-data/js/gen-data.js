/*require jszip.js FileSaver.js jquery*/

/* excel-gen.js

Client-Side JavaScript Code for creating Excel Spreadsheet tables from HTML Tables
Works on all browsers!!!!!

---------------
- MIT License -
---------------
Copyright 2018 ECSC, ltd.

Permission is hereby granted, free of charge, to any person obtaining a copy of this 
software and associated documentation files (the "Software"), to deal in the Software 
without restriction, including without limitation the rights to use, copy, modify, 
merge, publish, distribute, sublicense, and/or sell copies of the Software, and to 
permit persons to whom the Software is furnished to do so, subject to the following 
conditions:

The above copyright notice and this permission notice shall be included in all 
copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, 
INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A 
PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT 
HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF 
CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE 
OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.

Author Paul Warren */

//Initial XLSX Generation assumes default Workbook with sheet1, sheet2 and sheet3 inside.  Will upgrade to be more flexible in future releases.

/**
* Excel Generator.
*
* Creates .xlsx from HTML Table.
*
*/
function TDataGen(options) {
    //internal access to this
    var me = this;

    this.defaultOptions = {
        "src_id": "",
        "src": null,
        "format": "txt",
        "type": "table",
        "show_header": false,
        "auto_format": false,
        "header_row": null,
        "body_rows": null,
        "exclude_selector": null,
    }

    this.options = {};

    this.col_count = 0;
    this.columns = [];
    this.headers = [];
    this.rows = [];
    this.srcElem;

    this.range = "";

    /**** XML GENERATORS ****/

    /**
    * Creates sharedStrings.xml file.
    * 
    * Excel files have a sharedStrings.xml, this file holds all of the strings
    * used in the Excel spreadsheet to reduce repeating data.
    */
    this.sharedStrings = {
        "count": 0,
        "vals": [],
        /**
        * Adds value to Cache if it is a string and isn't already included.
        *
        * @returns {sharedString Value Object}
        */
        "add": function (value) {
            //update to specify format for input by column, or have autoformat.
            //based upon default excel format numbering system
            if (value.match(/^-?(?:\d+|\d{1,3}(?:,\d{3})+)(?:(\.|,)\d+)?$/)) {
                return { "type": "literal", "value": value.replace(/,/g,""), "text": value };
            } else if (me.__isDate__(value)) {
                var tmp = new Date(Date.parse(value));
                var ser = 25569.0 + ((tmp.getTime() - (tmp.getTimezoneOffset() * 60 * 1000)) / (1000 * 60 * 60 * 24))
                return { "type": "literal", "value": ser, "text": value };
            } else {
                this.count++;
                // value = me.encode(value);
                if (this.vals.indexOf(value) === -1) {
                    this.vals.push(value);
                }
                return { "type": "shared", "value": this.vals.indexOf(value), "text": value };
            }
        }
    };


    this.sheet = {
        "rows": []
    };

    String.prototype.format = function () {
        return (function (a, t) { return t.replace(/\{(\d+)\}/g, function (_, i) { return a[~ ~i] }) })(arguments, this);
    };

    this.__isDate__ = function (s) {
        // make sure it is in the expected format
        if (s.search(/^\d{1,2}[\/|\-|\.|_]\d{1,2}[\/|\-|\.|_]\d{4}/g) != 0)
            return false;

        // remove other separators that are not valid with the Date class    
        s = s.replace(/[\-|\.|_]/g, "/");

        // convert it into a date instance
        var dt = new Date(Date.parse(s));

        // check the components of the date
        // since Date instance automatically rolls over each component
        var arrDateParts = s.split("/");
        return (
             dt.getMonth() == arrDateParts[0] - 1 &&
             dt.getDate() == arrDateParts[1] &&
             dt.getFullYear() == arrDateParts[2]
         );
    }

    this.encode = function(str) {
        var hex = function (v) {
          return '&#x' + v.toString(16).toUpperCase() + ';';
        };
        
            var es = function(v) {
                return hex(v.charCodeAt(0));
            };
    
          str = str.replace(/["&'<>`]/g, es);
    
            return str.replace(/[\uD800-\uDBFF][\uDC00-\uDFFF]/g, function(v) {
                    var upper = v.charCodeAt(0), lower = v.charCodeAt(1), o = (upper - 0xD800) * 0x400 + lower - 0xDC00 + 0x10000;
                    return hex(o);
            }).replace(/[\x01-\t\x0B\f\x0E-\x1F\x7F\x81\x8D\x8F\x90\x9D\xA0-\uFFFF]/g, es);
    };

    /**
    *    Extension of JQuery Library for getting either the text or the value out of child elements.
    *
    *    Looks at the children elements, if it finds select or input tag, 
    *    it returns the value, otherwise it returns the text of the element.
    */
    jQuery.fn.extend({
        "textOrValue": function () {
            var t = this.find("select, input");
            return (t.length) ? t.val() : this.text();
        }
    });

    /**
    * Basic internal initialization.
    */
    this.__initialize__ = function (options) {
        this.options = $.extend(this.defaultOptions, options);
        this.__readHTMLTable__();
    };

    this.__readHTMLTable__ = function () {
        //setup HTML input
        if ((this.options.src_id) || (this.options.src)) {
            var table = (this.options.src_id) ? $("#" + this.options.src_id) : this.options.src;
            if ((table.length) && (table.prop("tagName") == "TABLE")) {
                var skipFirst = false;
                if ((!this.options.header_row) && (this.options.show_header)) {
                    if (table.has("thead").length) {
                        this.options.header_row = table.find("thead tr:nth-child(1)")
                    } else {
                        this.options.header_row = table.find("tr:nth-child(1)")
                        skipFirst = true;
                    }
                    this.col_count = this.options.header_row.length;
                }
                if (!this.options.body_rows) {
                    if (table.has("tbody").length) {
                        this.options.body_rows = (skipFirst) ? table.find("tbody tr").not(":first") : table.find("tbody tr");
                    } else {
                        this.options.body_rows = (skipFirst) ? table.find("tr").not(":first") : table.find("tr");
                    }
                    this.col_count = (this.col_count === 0) ? this.options.body_rows[0].length : this.col_count;
                }
            }
        }
        //process header if it exists
        if (this.options.header_row) {
            var row = [];
            var outerThis = this;
	        var colCount = 1;
            this.options.header_row.children("th,td").each(function () {
                var cell = $(this);
                if ((!outerThis.options.exclude_selector) || (cell.is(outerThis.options.exclude_selector) === false)) {
                    //header text gets stored for table
                    var txt = $(this).textOrValue().trim().replace(/ +(?= )/g, '');
                    if ((txt == "") && (outerThis.options.type == "table")) txt = "Column " + colCount;
                    if (txt!='選取'){//過濾帶有選取列
                        outerThis.headers.push(txt.replace(/[<]/g,""));
                        row.push(outerThis.sharedStrings.add(txt));
                        colCount++;
                    }
                }
            });
            this.sheet.rows.push(row);
        }
        //process content
        if (this.options.body_rows) {
            this.options.body_rows.each(function () {
                var row = [];
                $(this).children("th,td").each(function () {
                    var cell = $(this);
                    if ((!outerThis.options.exclude_selector) || (cell.is(outerThis.options.exclude_selector) === false)) {
                        var html=$(this).html();
                        if (html.indexOf("radio")<0 && html.indexOf("checkbox")<0){//過濾選取
                            row.push(outerThis.sharedStrings.add($(this).textOrValue().trim().replace(/ +(?= )/g, '')));
                        }
                    }
                });
                outerThis.sheet.rows.push(row);
            });
        }
    };


    this.generate = function () {
                var arrCSV = [];
                this.sheet.rows.forEach(function (r) {
                    var row = [];
                    r.forEach(function(c) {
                        var val = c.text.replace(/\"/g,"\"\"");
                        row.push(val+"\t");
                    })
                    arrCSV.push(row.join(" "));
                })
              var csv = arrCSV.join("\n");
            return csv;
     }



    //initialize the object
    this.__initialize__(options);

};
