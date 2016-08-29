/* Common app functionality */

var app = (function () {
    "use strict";

    var app = {};
    var host = 'https://uprinceuatapi.azurewebsites.net';
    // Common initialization function (to be called from each page)
    app.initialize = function () {
        $('body').append(
            '<div id="notification-message">' +
                '<div class="padding">' +
                    '<div id="notification-message-close"></div>' +
                    '<div id="notification-message-header"></div>' +
                    '<div id="notification-message-body"></div>' +
                '</div>' +
            '</div>');

        $('#notification-message-close').click(function () {
            $('#notification-message').hide();
        });


        // After initialization, expose a common notification function
        app.showNotification = function (header, text) {
            $('#notification-message-header').text(header);
            $('#notification-message-body').text(text);
            $('#notification-message').slideDown('fast');
        };

        function onBindingNotFound() {
            showMessage("The binding object was not found. " +
            "Please return to previous step to create the binding");
        }
        $(document).on("click", "#Refresh", function () {
            //deleteTable('ProductDescription');
            localStorage.setItem("dailyLog", "true");

            deleteTable('DailyLog');


        });

        $(document).on("click", "#Publish", function () {
            publishDailyLog();
        });

        $(document).on("click", "#createSheet", function () {
            Excel.run(function (ctx) {
                ctx.workbook.worksheets.add("Values");
                return ctx.sync().then(function () {
                    //app.showNotification("Success! Sheet created");
                    //riskRegisterGET();
                });
            }).catch(function (error) {
                //app.showNotification(error);
            });
            Excel.run(function (ctx) {
                ctx.workbook.worksheets.add("DailyLog");
                return ctx.sync().then(function () {
                    //app.showNotification("Success! Sheet created");
                    //riskRegisterGET();
                });
            }).catch(function (error) {
                //app.showNotification(error);
            });
            Excel.run(function (ctx) {
                ctx.workbook.worksheets.add("RiskRegister");
                return ctx.sync().then(function () {
                    //app.showNotification("Success! Sheet created");
                    //riskRegisterGET();
                });
            }).catch(function (error) {
                //app.showNotification(error);
            });
            Excel.run(function (ctx) {
                ctx.workbook.worksheets.add("ProductDescription");
                return ctx.sync().then(function () {
                    //app.showNotification("Success! Sheet created");
                    //riskRegisterGET();
                });
            }).catch(function (error) {
                //app.showNotification(error);
            });

        });
    };
    function isNull(param) {
        if (param == null) return '';
        else return param;
    }

    function isNull2(param) {
        if (param == "0" || param == "" || param == " ") return null;
        else return param
    }

    function isZero(param) {
        if (param == 0) return null
        else return param;
    }

    function isTypeSeverity(severity) {
        if (severity == "TeamManager") return "Team Manager";
        else if (severity == "ProjectManager") return "Project Manager";
        else if (severity == "ProjectBoard") return "Project Board";
        else if (severity == "CooperateProgramManagement") return "Corporate / Program Management";
    }

    //if date is given as a regular date yyyy-mm-ddT..
    function formatDate(date) {
        if (date == null || date == "") { return '' }
            //app.showNotification();
        else {
            date = date.substring(0, 10);
            //app.showNotification(date);
            var year = date.substring(0, 4);
            var month = date.substring(5, 7);
            if (month.charAt(0) == "0") month = month.substring(1, 2);
            var day = date.substring(8, 10)
            if (day.charAt(0) == "0") day = day.substring(1, 2);
            //app.showNotification(year + "-" + month + "-" + day)
            return year + "-" + month + "-" + day;
        };
    }

    //if date is given in second since 01-01-1970
    function formatDate2(dateS) {
        if (dateS == null) return '';
        else {
            var mSeconds = dateS * 1000;
            var date = new Date(mSeconds);
            var day = date.getDate();
            var month = date.getMonth() + 1;
            var year = date.getFullYear();
            //app.showNotification(year + "-" + month + "-" + day);
            return year + "-" + month + "-" + day;
        };
    };

    function formatDate3(date2) {
        if (date2 == null || date2 == 0) return null;
        else {
            var date = new Date((date2 - (25567 + 2)) * 86400 * 1000);

            return formatDate(date.toJSON());
        }
    };

    //if date is asked in form yyyy-mm-dd
    function convertDate(days) {
        if (days == "" || days == null) { return null }
        else {
            var dateDays = days - 25569;
            var dateMS = dateDays * 24 * 60 * 60 * 1000;
            var date = new Date(dateMS);
            var day = date.getDate();
            var month = date.getMonth() + 1;
            var year = date.getFullYear();
            return (year + "-" + month + "-" + day + "T00:00:00.000");
        }
    }

    //if date is asked in seconds since 01-01-1970
    function convertDate2(days) {
        if (days == "" || days == null) { return null }
        else {
            var dateDays = days - 25569;
            var dateS = dateDays * 24 * 60 * 60;
            return dateS
        }
    }

    function activateWorksheet(name) {
        Excel.run(function (ctx) {
            var wSheetName = name;
            var worksheet = ctx.workbook.worksheets.getItem(wSheetName);
            worksheet.activate();
            return ctx.sync();
        })
    .catch(function (error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    });
    };

    function deleteTable(name) {
        Excel.run(function (ctx) {
            var tableName = 'DailyLog';
            var tableName2 = 'Projects';
            var tableName3 = 'atContext';
            var tableName4 = 'Status';
            var tableName5 = 'Responsible';
            var tableName6 = 'Type';
            var tableName7 = 'Time';
            var tableName8 = 'Energy';
            var table = ctx.workbook.tables.getItem(tableName);
            var table2 = ctx.workbook.tables.getItem(tableName2);
            var table3 = ctx.workbook.tables.getItem(tableName3);
            var table4 = ctx.workbook.tables.getItem(tableName4);
            var table5 = ctx.workbook.tables.getItem(tableName5);
            var table6 = ctx.workbook.tables.getItem(tableName6);
            var table7 = ctx.workbook.tables.getItem(tableName7);
            var table8 = ctx.workbook.tables.getItem(tableName8);
            var tableDataRange = table.getDataBodyRange();
            var tableDataRange2 = table2.getDataBodyRange();
            var tableDataRange3 = table3.getDataBodyRange();
            var tableDataRange4 = table4.getDataBodyRange();
            var tableDataRange5 = table5.getDataBodyRange();
            var tableDataRange6 = table6.getDataBodyRange();
            var tableDataRange7 = table7.getDataBodyRange();
            var tableDataRange8 = table8.getDataBodyRange();
            tableDataRange.load('address');
            tableDataRange2.load('address');
            tableDataRange3.load('address');
            tableDataRange4.load('address');
            tableDataRange5.load('address');
            tableDataRange6.load('address');
            tableDataRange7.load('address');
            tableDataRange8.load('address');
            //var range = tableDataRange.address;
            //tableDataRange.address.delete();
            return ctx.sync().then(function () {
                //console.log(tableDataRange.address);
                var range = tableDataRange.address;
                var range2 = tableDataRange2.address;
                var range3 = tableDataRange3.address;
                var range4 = tableDataRange4.address;
                var range5 = tableDataRange5.address;
                var range6 = tableDataRange6.address;
                var range7 = tableDataRange7.address;
                var range8 = tableDataRange8.address;
                var rangeAddress = range.substring(range.indexOf('!') + 1);
                var rangeAddress2 = range2.substring(range2.indexOf('!') + 1);
                var rangeAddress3 = range3.substring(range3.indexOf('!') + 1);
                var rangeAddress4 = range4.substring(range4.indexOf('!') + 1);
                var rangeAddress5 = range5.substring(range5.indexOf('!') + 1);
                var rangeAddress6 = range6.substring(range6.indexOf('!') + 1);
                var rangeAddress7 = range7.substring(range7.indexOf('!') + 1);
                var rangeAddress8 = range8.substring(range8.indexOf('!') + 1);
                localStorage.setItem('rangeAddress', rangeAddress);
                localStorage.setItem('rangeAddress2', rangeAddress2);
                localStorage.setItem('rangeAddress3', rangeAddress3);
                localStorage.setItem('rangeAddress4', rangeAddress4);
                localStorage.setItem('rangeAddress5', rangeAddress5);
                localStorage.setItem('rangeAddress6', rangeAddress6);
                localStorage.setItem('rangeAddress7', rangeAddress7);
                localStorage.setItem('rangeAddress8', rangeAddress8);

                var sheetName = range.substring(0, range.indexOf('!'));
                var sheetName2 = range2.substring(0, range2.indexOf('!'));
                var sheetName3 = range3.substring(0, range3.indexOf('!'));
                var sheetName4 = range4.substring(0, range4.indexOf('!'));
                var sheetName5 = range5.substring(0, range5.indexOf('!'));
                var sheetName6 = range6.substring(0, range6.indexOf('!'));
                var sheetName7 = range7.substring(0, range7.indexOf('!'));
                var sheetName8 = range8.substring(0, range8.indexOf('!'));
                localStorage.setItem('sheetName', sheetName);
                localStorage.setItem('sheetName2', sheetName2);
                localStorage.setItem('sheetName3', sheetName3);
                localStorage.setItem('sheetName4', sheetName4);
                localStorage.setItem('sheetName5', sheetName5);
                localStorage.setItem('sheetName6', sheetName6);
                localStorage.setItem('sheetName7', sheetName7);
                localStorage.setItem('sheetName8', sheetName8);
            });
        }).catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });

        Excel.run(function (ctx) {
            var range = ctx.workbook.worksheets.getItem(localStorage.getItem("sheetName")).getRange(localStorage.getItem("rangeAddress"));
            var range2 = ctx.workbook.worksheets.getItem(localStorage.getItem("sheetName2")).getRange(localStorage.getItem("rangeAddress2"));
            var range3 = ctx.workbook.worksheets.getItem(localStorage.getItem("sheetName3")).getRange(localStorage.getItem("rangeAddress3"));
            var range4 = ctx.workbook.worksheets.getItem(localStorage.getItem("sheetName4")).getRange(localStorage.getItem("rangeAddress4"));
            var range5 = ctx.workbook.worksheets.getItem(localStorage.getItem("sheetName5")).getRange(localStorage.getItem("rangeAddress5"));
            var range6 = ctx.workbook.worksheets.getItem(localStorage.getItem("sheetName6")).getRange(localStorage.getItem("rangeAddress6"));
            var range7 = ctx.workbook.worksheets.getItem(localStorage.getItem("sheetName7")).getRange(localStorage.getItem("rangeAddress7"));
            var range8 = ctx.workbook.worksheets.getItem(localStorage.getItem("sheetName8")).getRange(localStorage.getItem("rangeAddress8"));
            //app.showNotification(localStorage.getItem("rangeAddress"));
            range.delete();
            range2.delete();
            range3.delete();
            range4.delete();
            range5.delete();
            range6.delete();
            range7.delete();
            range8.delete();
            return ctx.sync().then(function () {
                //dailyLogGET();
            });
        }).catch(function (error) {
            //app.showNotification("Error: " + error);
            //dailyLogGET();
            if (error instanceof OfficeExtension.Error) {
                //app.showNotification("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
        //dailyLogGET();
        Excel.run(function (ctx) {
            var tableName = name;
            var table = ctx.workbook.tables.getItem(tableName);
            var tableDataRange = table.getDataBodyRange();
            tableDataRange.load('address')
            //var range = tableDataRange.address;
            //tableDataRange.address.delete();
            return ctx.sync().then(function () {
                //console.log(tableDataRange.address);
                var range = tableDataRange.address;
                var rangeAddress = range.substring(range.indexOf('!') + 1);
                //var rowCounts = rangeAddress.substring(1);
                //app.showNotification(rangeAddress)
                if (rangeAddress != "A2:K2" && ctx.workbook.tables.getItem("DailyLog").rows.load("values").items[0].values[0][1] != '') {
                    deleteTable(name);
                    app.showNotification(ctx.workbook.tables.getItem("DailyLog").rows.load("values").items[0].values[0][1]);
                }
                else if (localStorage.getItem("dailyLog") == "true") {
                    dailyLogGET();
                }
            });
        }).catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    };

    //Daily Log
    function dailyLogGET() {
        localStorage.setItem("dailyLog", "false")
        //deleteTable("DailyLog");
        //var projectId = localStorage.getItem('projectId');
        var userEmail = localStorage.getItem('email');
        var urlProject = host + "/api/DailyLog/GetDailyLog";
        var dataEmail =
            {
                "projectId": "",
                "project": null,
                "identifier": "",
                "title": "",
                "activity": "",
                "responsibleStatusString": "",
                "status": {
                    "All": true,
                    "New": false,
                    "Waiting": false,
                    "Completed": false
                },
                "activityType": {
                    "All": true,
                    "Problem": false,
                    "Action": false,
                    "Event": false,
                    "Comment": false,
                    "Decision": false,
                    "Reference": false
                },
                "priority": {
                    "All": true,
                    "High": false,
                    "Medium": false,
                    "Low": false
                },
                "targetDate": {
                    "All": true,
                    "-7": false,
                    "-1": false,
                    "0": false,
                    "+1": false,
                    "+7": false
                },
                "responsibleStatus": {
                    "All": true,
                    "Inbox": false,
                    "Next": false,
                    "Waiting": false,
                    "Schedule": false,
                    "Someday": false,
                    "Done": false
                },
                "energy": {
                    "All": true,
                    "Mild": false,
                    "Reasonable": false,
                    "Demanding": false,
                    "Very Demanding": false,
                    "Extreme": false
                },
                "responsible": "",
                "requester": "",
                "coreUserEmail": userEmail,
                "atContext": "",
                "startDate": "",
                "orderField": "",
                "sortOrder": ""
            };
        $.ajax({
            type: "POST",
            url: urlProject,
            dataType: "json",
            contentType: "application/json; charset=utf-8",
            data: JSON.stringify(dataEmail),
        })
            .done(function (str) {
                //app.showNotification(str.dailyLogListViewModel.length);
                getDailyLog(str.dailyLogListViewModel[1].id);
                var DlId = [str.dailyLogListViewModel.length];
                if (str.dailyLogListViewModel.length > 0) {
                    var matrix = [str.dailyLogListViewModel.length];
                    for (var i = 0; i < str.dailyLogListViewModel.length; i++) {
                        matrix[i] = [11];
                        DlId[i] = str.dailyLogListViewModel[i].id;
                        matrix[i][0] = isNull(str.dailyLogListViewModel[i].project);
                        var activity = isNull(str.dailyLogListViewModel[i].activity);
                        matrix[i][1] =activity;
                        var identifier = isNull(str.dailyLogListViewModel[i].id);
                        matrix[i][10] = identifier;
                        matrix[i][3] = isNull(str.dailyLogListViewModel[i].atContext);
                        //matrix[i][4] = str.dailyLogListViewModel[i].targetDate;
                        matrix[i][4] = formatDate(str.dailyLogListViewModel[i].targetDate);
                        matrix[i][5] = isNull(str.dailyLogListViewModel[i].responsibleStatus);
                        matrix[i][6] = isNull(str.dailyLogListViewModel[i].responsible);
                        matrix[i][7] = isNull(str.dailyLogListViewModel[i].activityType);
                        matrix[i][8] = isNull(str.dailyLogListViewModel[i].time);
                        matrix[i][9] = isNull(str.dailyLogListViewModel[i].energy);
                        matrix[i][2] = '=HYPERLINK("https://uprinceuat.azurewebsites.net/Home/Index/gdt/daily-log/description/"&[@Column1],[@Column1])'
                        //matrix[i][10] = 'hello';
                        storeLineDaily(matrix[i]);
                        //GTD > excel > Hyperlink 

                        //app.showNotification(matrix[i].toString())
                        //matrix[i][6] = isNull(str[i].Version);
                        //localStorage.setItem("ParentId" + str[i].Id, str[i].ParentId);
                    }
                } else {
                    var matrix = [["", "", "", "", "", "", "", "", "", "",""]]
                };

                localStorage.setItem("DlId", DlId);

                Excel.run(function (ctx) {
                    var tables = ctx.workbook.tables;
                    var tableRows = tables.getItem('DailyLog').rows
                    for (var i = 0; i < matrix.length; i++) {
                        var line = [1];
                        line[0] = matrix[i];
                        tableRows.add(null, line);
                    };
                    return ctx.sync().then(function () {
                        showMessage("Success! My monthly expense table created! Select the arrow button to see how to remove the table.");
                        storeLineDaily(matrix[i]);
                    })
                     .catch(function (error) {
                         showMessage(JSON.stringify(error));
                     });
                });


            });
    };

    function addValues(tableName, row, ctx) {
        var tables = ctx.workbook.tables;
        var tableRows = tables.getItem(tableName).rows
        for (var i = 0; i < row.length; i++) {
            var line = [1];
            line[0] = row[i];
            tableRows.add(null, line);
        }
    };

    function getDailyLog(dailyLogId) {
        var urlProject = host + '/api/DailyLog/GetDailyLog?logId=' + dailyLogId + '&email=' + localStorage.getItem("email");
        $.ajax({
            type: 'GET',
            url: urlProject,
            dataType: "json",
            contentType: "application/json; charset=utf-8",

        })
         .done(function (str) {
             //app.showNotification(str.impact[0].State);
             Excel.run(function (ctx) {
                 //var matrix = riskValuesImpact(str);
                 if (Object.keys(str.contextList).length != 0) {
                     //ctx.workbook.worksheets.getItem('Values').getRange("G1:G" + Object.keys(str.contextList).length).values = dailyLogContext(str);
                     addValues("atContext", dailyLogContext(str), ctx);
                 }
                 if (Object.keys(str.personnelContacts).length != 0) {
                     addValues("Responsible", dailyLogUsers(str), ctx);
                     //ctx.workbook.worksheets.getItem('Values').getRange("I1:I" + Object.keys(str.personnelContacts).length).values = dailyLogUsers(str)
                 }
                 var status = [["Inbox"], ["Next"], ["Waiting"], ["Schedule"], ["Someday"], ["Done"]];
                 addValues("Status", status, ctx);
                 var type = [["Problem"], ["Action"], ["Event"], ["Comment"], ["Decision"], ["Reference"]];
                 addValues("Type", type, ctx);
                 var time = [["5 min"], ["15 min"], ["30 min"], ["1 hr"], ["2 hr"], ["4 hr"], ["8 hr"],["' -"]]/*[[1], [2], [1], [2], [1]] //str.impact[0].State*/;
                 addValues("Time", time, ctx);
                 var energy = [["Mild"], ["Reasonable"], ["Demanding"], ["Very Demanding"], ["Extreme"],["' -"]]/*[[1], [2], [1], [2], [1]] //str.impact[0].State*/;
                 addValues("Energy", energy, ctx);

                 if (Object.keys(str.project).length != 0) {
                     addValues("Projects", dailyLogProject(str), ctx)
                     //ctx.workbook.worksheets.getItem('Values').getRange("M1:M" + Object.keys(str.project).length).values = dailyLogProject(str);
                 }

                 return ctx.sync().then(function () {
                     //console.log("Success! Insert range in A1:C3.");
                 });
             }).catch(function (error) {
                 //app.showNotification(error);
             });
         })
    };

    function dailyLogContext(str) {
        var val = [Object.keys(str.contextList).length];
        for (var i = 0; i < Object.keys(str.contextList).length; i++) {
            val[i] = [1];
            val[i][0] = str.contextList[i].description;
            localStorage.setItem('dailyLogContext' + str.contextList[i].description, "" + str.contextList[i].id);
            //val[i] = str.impact[i].State;
        }
        //app.showNotification(val[2][0]);
        return val;
    };

    function dailyLogUsers(str) {
        var val = [Object.keys(str.personnelContacts).length];
        for (var i = 0; i < Object.keys(str.personnelContacts).length; i++) {
            val[i] = [1];
            val[i][0] = str.personnelContacts[i].State;
            localStorage.setItem('dailyLogUsers' + str.personnelContacts[i].State, "" + str.personnelContacts[i].StateId);
            //val[i] = str.impact[i].State;
        }
        //app.showNotification(val[2][0]);
        return val;
    }

    function dailyLogProject(str) {
        var val = [Object.keys(str.project).length];
        for (var i = 0; i < Object.keys(str.project).length; i++) {
            val[i] = [1];
            val[i][0] = str.project[i].State;
            localStorage.setItem('dailyLogProject' + str.project[i].State, "" + str.project[i].StateId);
            //val[i] = str.impact[i].State;
        }
        //app.showNotification(val[2][0]);
        return val;
    }

    function publishDailyLog() {
        Excel.run(function (ctx) {
            var rows = ctx.workbook.tables.getItem("DailyLog").rows.load("values");
            return ctx.sync()
                .then(function () {
                    var DlId = localStorage.getItem('DlId');
                    var projectId = localStorage.getItem('projectId');
                    for (var i = 0; i < rows.items.length; i++) {
                        var row = rows.items[i];
                        if (isNewDaily(row)) {
                            commitNewDaily(row, i);
                        }
                        else {
                            if (editedLine(row)) {
                                commitUpdateDaily(row);
                            }
                        };
                    }
                })
                .then(ctx.sync)
                .then(function () {
                    console.log("Success! Format rows of 'Table1' with 2nd cell greater than 2 in green, other rows in red.");
                });
        }).catch(function (error) {
            console.log(error);
        });
    };

    function addId(range, id) {
        Excel.run(function (ctx) {
            var count = (range + 2);
            //app.showNotification(typeof(range) + "  "+ typeof(range + 1));
            var excelRange = "K" + count + ":K" + count;
            //app.showNotification(excelRange);
            ctx.workbook.worksheets.getItem('DailyLog').getRange(excelRange).values = ("" + id)/*[[1], [2], [1], [2], [1]] //str.impact[0].State*/;

            return ctx.sync().then(function () {
                //console.log("Success! Insert range in A1:C3.");
            });
        }).catch(function (error) {
            // app.showNotification(error);
        });
    };

    function isActivityType(type) {
        if (type == "Problem") return "0";
        else if (type == "Action") return "1";
        else if (type == "Event") return "2";
        else if (type == "Comment") return "3";
        else if (type == "Decision") return "4";
        else if (type == "Reference") return "5";
        else return null
    }

    function isEnergy(type) {
        if (type == "Mild") return "0";
        else if (type == "Reasonable") return "1";
        else if (type == "Demanding") return "2";
        else if (type == "Very Demanding") return "3";
        else if (type == "Extreme") return "4";
        else return null;
    };

    function isTime(type) {
        if (type == "5 min") return "0";
        else if (type == "15 min") return "1";
        else if (type == "30 min") return "2";
        else if (type == "1 hr") return "3";
        else if (type == "2 hr") return "4";
        else if (type == "4 hr") return "5";
        else if (type == "8 hr") return "6";
        else return null;
    };

    function isResponseStatus(type) {
        if (type == "Inbox") return "0";
        else if (type == "Next") return "1";
        else if (type == "Waiting") return "2";
        else if (type == "Schedule") return "3";
        else if (type == "Someday") return "4";
        else if (type == "Done") return "5";
        else return null;
    };

    function commitNewDaily(line, i) {
        var email = localStorage.getItem("email");
        var urlProject = host + '/api/DailyLog/PostDailyLogHeader';
        var urlProject2 = host + '/api/DailyLog/PostDailyLogInvolvedTiming';
        var dataEmail = {
            "id": isNull2(line.values[0][10]),
            "projectId": projectId(line.values[0][0], line, 'commitNew', i),
            "activityTypeId": isActivityType(line.values[0][7]),
            "title": isNull(line.values[0][1]),
            "order": "0",
            "coreUserEmail": email,
            "authorEmail": email,
            "context": contextId(line.values[0][3], line, 'commitNew', i),
            "energy": isEnergy(line.values[0][9])
        };
        $.ajax({
            type: "POST",
            url: urlProject,
            dataType: "json",
            async: false,
            contentType: "application/json; charset=utf-8",
            data: JSON.stringify(dataEmail),
        })
        .done(function (result) {
            addId(i, result);
            var dataEmail2 = {
                "dailyLogId": result,
                "responsible": localStorage.getItem('dailyLogUsers' + line.values[0][6]),
                "responseStatus": isResponseStatus(line.values[0][5]),
                "targetDate": formatDate3(line.values[0][4]),
                "time": isTime(line.values[0][8])
            };
            $.ajax({
                type: "POST",
                url: urlProject2,
                dataType: "json",
                contentType: "application/json; charset=utf-8",
                data: JSON.stringify(dataEmail2),
            });
        })
    };

    function commitUpdateDaily(line) {
        //app.showNotification(line.values[0].toString());
        var urlProject = host + '/api/DailyLog/PostDailyLogHeader';
        var urlProject2 = host + '/api/DailyLog/PostDailyLogInvolvedTiming';
        var dataEmail = {
            "id": isNull2(line.values[0][10]),
            "projectId": projectId(line.values[0][0], line, 'commitUpdate', ""),
            "activityTypeId": isActivityType(line.values[0][7]),
            "title": isNull(line.values[0][1]),
            "order": "0",
            "coreUserEmail": localStorage.getItem("email"),
            "authorEmail": localStorage.getItem("email"),
            "context": contextId(line.values[0][3], line, 'commitUpdate', ""),
            "energy": isEnergy(line.values[0][9])
        };
        $.ajax({
            type: "POST",
            url: urlProject,
            dataType: "json",
            contentType: "application/json; charset=utf-8",
            data: JSON.stringify(dataEmail),
        })
        var dataEmail2 = {
            "dailyLogId": isNull2(line.values[0][10]),
            "responsible": localStorage.getItem('dailyLogUsers' + line.values[0][6]),
            "responseStatus": isResponseStatus(line.values[0][5]),
            "targetDate": formatDate3(line.values[0][4]),
            "time": isTime(line.values[0][8])
        };
        //app.showNotification(line.values[0][4]);
        //line.values[0][4] = "";
        $.ajax({
            type: "POST",
            url: urlProject2,
            dataType: "json",
            contentType: "application/json; charset=utf-8",
            data: JSON.stringify(dataEmail2),
        });
        //app.showNotification(line.values[0].toString());
    };

    function isNewDaily(line) {
        if (line.values[0][10] == null || line.values[0][10] == "") return true;
        else return false;
    }

    function storeLineDaily(line) {
        //line[4] = formatDate(line[4]);
        localStorage.setItem("line" + line[2], line.toString());
        //app.showNotification(line.toString());
        //app.showNotification(line[4]);
    }

    function editedLine(line) {
        var oldLine = localStorage.getItem("line" + line.values[0][10]);
        var temporary = line.values[0][4];
        line.values[0][4] = formatDate3(line.values[0][4]);
        //app.showNotification(oldLine + " / " + line.values[0].toString());
        if (line.values[0].toString() != oldLine) {
            line.values[0][4] = temporary;
            return true;
        }
        else {
            line.values[0][4] = temporary;
            return false
        };
    }

    function projectId(project, line, type, i) {
        //app.showNotification(localStorage.getItem("dailyLogProject" + project) == null);
        if (localStorage.getItem("dailyLogProject" + project) == null) {
            var dataEmail = {
                "id": null,
                "title": project,
                "status": "1",
                "statusStyle": "green-border",
                "statusIcon": "green-icon.png",
                "projectOwnerId": localStorage.getItem("uId"),
                "pmEmail": localStorage.getItem("email"),
                "customer": "",
                "startDate": null,
                "expectedEndDate": null,
                "endDate": null,
                "estimatedBudget": null,
                "budget": null,
                "toleranceStatus": null,
                "selectedToleranceStatus": null,
                "workflowStatus": null,
                "projectBusinessId": "",
                "selectedProjectTypeId": null,
                "selectedProjectManagementLevelId": 3,
                "isPublic": false,
                "createdDateTime": null
            };
            var urlProject = host + "/api/project/PostProject";
            $.ajax({
                type: "POST",
                url: urlProject,
                async: false,
                dataType: "json",
                contentType: "application/json; charset=utf-8",
                data: JSON.stringify(dataEmail),
            })
            .done(function (result) {
                localStorage.setItem("dailyLogProject" + project, "" + result);
                app.showNotification("" + result);
                if (type == "commitUpdate") { commitUpdateDaily(line) }
                else commitNewDaily(line, i)
                return "" + result;
            })
        }
        else return localStorage.getItem("dailyLogProject" + project);
    };

    function contextId(context, line, type, i) {
        var userEmail = localStorage.getItem("email");
        //app.showNotification(localStorage.getItem("dailyLogProject" + project) == null);
        if (localStorage.getItem("dailyLogContext" + context) == null) {
            var dataEmail = {
                "id": null,
                "description": context,
                "coreUserEmail": userEmail,
            
            };
            var urlProject = host + "/api/DailyLog/PostDailyLogContext";
            $.ajax({
                type: "POST",
                url: urlProject,
                async: false,
                dataType: "json",
                contentType: "application/json; charset=utf-8",
                data: JSON.stringify(dataEmail),
            })
            .done(function (result) {
                localStorage.setItem("dailyLogContext" + context, "" + result);
                app.showNotification("New @Context added");

                if (type == "commitUpdate") {
                    commitUpdateDaily(line);
                    commitUpdateDaily(line);
                }
                else commitNewDaily(line, i);
                app.showNotification(typeof (""+result))
                return ("" + result);
            })
        }
        else return localStorage.getItem("dailyLogContext" + context);
    };
    return app;
})();