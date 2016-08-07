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
            //deleteTable('ProductDescription');
            //deleteTable('risk)
            //riskRegisterGET();
            //productDescriptionGET();
            //dailyLogGET();
        });

        $(document).on("click", "#Publish", function () {
            //publishRiskRegister();
            //publishProductDescription();
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
        else return date.substring(0, 10);
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

    function formatDate3(date) {
        if (date == null || date == 0) return null;
        else {
            var date = new Date((date - (25567 + 2)) * 86400 * 1000);
            return date.getFullYear + "-" + (date.getMonth + 1) + "-" + date.getDate;
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
                localStorage.setItem('rangeAddress', rangeAddress);
                var sheetName = range.substring(0, range.indexOf('!'));
                localStorage.setItem('sheetName', sheetName);
                localStorage.setItem("rowCounts", rowCounts);
            });
        }).catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
        //var length = parseInt(localStorage.getItem('rangeAddress'))-1;
        //app.showNotification(length);
        /*for (var i = 0; i < length; i++) {
            Excel.run(function (ctx) {
                ctx.workbook.tables.getItem(name).rows.getItemAt(0).delete();
                return ctx.sync().then(function () {
                    console.log("Success! Deleted the 0th row from 'MyTable' and shifts cells up.");
                });;
            }).catch(function (error) {
                console.log("Error: " + error);
            });
        }*/
        Excel.run(function (ctx) {
            var range = ctx.workbook.worksheets.getItem(localStorage.getItem("sheetName")).getRange(localStorage.getItem("rangeAddress"));
            //app.showNotification(localStorage.getItem("rangeAddress"));
            range.delete();
            return ctx.sync();
        }).catch(function (error) {
            //app.showNotification("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                //app.showNotification("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });

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
                if (rangeAddress != "A2:J2" && ctx.workbook.tables.getItem("DailyLog").rows.load("values").items[0].values[0][1] != '') {
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
        var projectId = localStorage.getItem('projectId');
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
                getDailyLog(str.dailyLogListViewModel[1].identifier);
                var DlId = [str.dailyLogListViewModel.length];
                if (str.dailyLogListViewModel.length > 0) {
                    var matrix = [str.dailyLogListViewModel.length];
                    for (var i = 0; i < str.dailyLogListViewModel.length; i++) {
                        matrix[i] = [10];
                        DlId[i] = str.dailyLogListViewModel[i].identifier;
                        matrix[i][0] = isNull(str.dailyLogListViewModel[i].project);
                        matrix[i][1] = isNull(str.dailyLogListViewModel[i].activity);
                        matrix[i][2] = isNull(str.dailyLogListViewModel[i].identifier);
                        matrix[i][3] = isNull(str.dailyLogListViewModel[i].atContext);
                        matrix[i][4] = formatDate(str.dailyLogListViewModel[i].targetDate);
                        matrix[i][5] = isNull(str.dailyLogListViewModel[i].responsibleStatus);
                        matrix[i][6] = isNull(str.dailyLogListViewModel[i].responsible);
                        matrix[i][7] = isNull(str.dailyLogListViewModel[i].activityType);
                        matrix[i][8] = isNull(str.dailyLogListViewModel[i].time);
                        matrix[i][9] = isNull(str.dailyLogListViewModel[i].energy);
                        //app.showNotification(matrix[i].toString())
                        //matrix[i][6] = isNull(str[i].Version);
                        //localStorage.setItem("ParentId" + str[i].Id, str[i].ParentId);
                    }
                } else {
                    var matrix = [["", "", "", "", "", "", "", "", "", ""]]
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
                     ctx.workbook.worksheets.getItem('Values').getRange("G1:G" + Object.keys(str.contextList).length).values = dailyLogContext(str)/*[[1], [2], [1], [2], [1]] //str.impact[0].State*/;
                 }
                 if (Object.keys(str.personnelContacts).length != 0) {
                     ctx.workbook.worksheets.getItem('Values').getRange("I1:I" + Object.keys(str.personnelContacts).length).values = dailyLogUsers(str)
                 }
                 ctx.workbook.worksheets.getItem('Values').getRange("H1:H6").values = [["Inbox"], ["Next"], ["Waiting"], ["Schedule"], ["Someday"], ["Done"]]/*[[1], [2], [1], [2], [1]] //str.impact[0].State*/;
                 ctx.workbook.worksheets.getItem('Values').getRange("J1:J6").values = [["Problem"], ["Action"], ["Event"], ["Comment"], ["Decision"], ["Reference"]]/*[[1], [2], [1], [2], [1]] //str.impact[0].State*/;
                 ctx.workbook.worksheets.getItem('Values').getRange("K1:K7").values = [["5 min"], ["15 min"], ["30 min"], ["1 hr"], ["2 hr"], ["4 hr"], ["8 hr"]]/*[[1], [2], [1], [2], [1]] //str.impact[0].State*/;
                 ctx.workbook.worksheets.getItem('Values').getRange("L1:L5").values = [["Mild"], ["Reasonable"], ["Demanding"], ["Very Demanding"], ["Extreme"]]/*[[1], [2], [1], [2], [1]] //str.impact[0].State*/;
                 if (Object.keys(str.project).length != 0) {
                     ctx.workbook.worksheets.getItem('Values').getRange("M1:M" + Object.keys(str.project).length).values = dailyLogProject(str);
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
            var excelRange = "C" + count + ":C" + count;
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
            "id": line.values[0][2],
            "projectId": localStorage.getItem("dailyLogProject" + line.values[0][0]),
            "activityTypeId": isActivityType(line.values[0][7]),
            "title": isNull(line.values[0][1]),
            "order": "0",
            "coreUserEmail": email,
            "authorEmail": email,
            "context": localStorage.getItem("dailyLogContext" + line.values[0][3]),
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
            "id": line.values[0][2],
            "projectId": localStorage.getItem("dailyLogProject" + line.values[0][0]),
            "activityTypeId": isActivityType(line.values[0][7]),
            "title": isNull(line.values[0][1]),
            "order": "0",
            "coreUserEmail": localStorage.getItem("email"),
            "authorEmail": localStorage.getItem("email"),
            "context": localStorage.getItem("dailyLogContext" + line.values[0][3]),
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
            "dailyLogId": line.values[0][2],
            "responsible": localStorage.getItem('dailyLogUsers' + line.values[0][6]),
            "responseStatus": isResponseStatus(line.values[0][5]),
            "targetDate": formatDate3(line.values[0][4]),
            "time": isTime(line.values[0][8])
        };
        line.values[0][4] = "";
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
        if (line.values[0][2] == null || line.values[0][2] == "") return true;
        else return false;
    }

    function storeLineDaily(line) {
        line[4] = "";
        localStorage.setItem("line" + line[2], line.toString());
        //app.showNotification(line.toString());
    }

    function editedLine(line) {
        var oldLine = localStorage.getItem("line" + line.values[0][2]);
        app.showNotification(line.values[0][4]);
        line.values[0][4] = "";
        //app.showNotification(oldLine + " / " + line.values[0].toString());
        if (line.values[0].toString() != oldLine) {
            //app.showNotification("edited")
            return true;
        }
        else return false;
    }
    return app;
})();