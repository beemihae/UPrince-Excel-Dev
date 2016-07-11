﻿/* Common app functionality */

var app = (function () {
    "use strict";

    var app = {};
    var host = 'https://uprincecoredevapi.azurewebsites.net';
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

        $(document).on("click", '#Cmdb', function () {
            //id excluded
            var projectId = sessionStorage.getItem('projectId'); //when using the login-screen
            //app.showNotification(projectId);

            //var projectId = '22050'; //to test just this page
            var urlProject = host + '/api/ConfigManagerDb/GetConfigurationManagerRegister';
            var dataEmail = {
                "projectId": projectId,
                "sortField": "title",
                "sortOrder": "DESC",
                "status": {
                    "All": "true",
                    "PendingDevelopment": "false",
                    "InDevelopment": "false",
                    "InReview": "false",
                    "Approved": "false",
                    "HandedOver": "false"
                },
                "type": {
                    "All": "true",
                    "Component": "false",
                    "Product": "false",
                    "Release": "true"
                },
                "title": "",
                "identifier": ""
            }


            $.ajax({
                type: "POST",
                url: host + "/api/ConfigManagerDb/GetConfigurationManagerRegister",
                dataType: "json",
                contentType: "application/json; charset=utf-8",
                data: JSON.stringify(dataEmail),
            })
              .done(function (str) {
                  var CmdbId = [str.length];
                  var length = Object.keys(str).length;
                  if (length > 0) {
                      var matrix = [length];
                      for (var i = 0; i < length; i++) {
                          matrix[i] = [6];
                          CmdbId[i] = str[i].id;
                          matrix[i][0] = isNull(str[i].title);
                          matrix[i][1] = isNull(str[i].identifier);
                          matrix[i][2] = isNull(str[i].status);
                          matrix[i][3] = isNull(str[i].type);
                          matrix[i][4] = isNull(str[i].location);
                          matrix[i][5] = isNull(str[i].producer);
                      }
                  }
                  else {
                      var matrix = [["", "", "", "", "", ""]]
                  }
                  sessionStorage.setItem("CmdbId", CmdbId);
                  var cmdb = new Office.TableData();
                  cmdb.headers = ["Title", "ID", "Status", "Type", "Location", "Producer"];
                  cmdb.rows = matrix;
                  // Set the myTable in the document.
                  Office.context.document.setSelectedDataAsync(
                    cmdb,
                    {
                        coercionType: Office.CoercionType.Table, cellFormat: [{ cells: Office.Table.All, format: { width: "auto fit" } }
                        ]
                    },
                    function (asyncResult) {
                        if (asyncResult.status == "failed") {
                            //showMessage("Action failed with error: " + asyncResult.error.message);
                        } else {
                            //showMessage("Check out your new table, then click next to learn another API call.");
                        }
                    }
                  );

                  Office.context.document.bindings.addFromSelectionAsync(
                       Office.BindingType.Table,
                               { id: "cmdb" },
                       function (asyncResult) {
                           if (asyncResult.status == "failed") {
                               //showMessage("Action failed with error: " + asyncResult.error.message);
                           } else {
                               //app.showNotification('Binding done');
                           }
                       });
              });
        });

        $(document).on("click", '#DailyLog', function () {
            var projectId = sessionStorage.getItem('projectId'); //when using the login-screen
            //app.showNotification(projectId);
            app.showNotification('Hello');
            //var projectId = '22050'; //to test just this page
            var urlProject = host + '/api/DailyLog/GetDailyLog';
            var dataEmail = {
                "projectId": projectId,
                "identifier": "",
                "title": "",
                "status": {
                    "All": true,
                    "New": false,
                    "Waiting": false,
                    "Completed": false
                },
                "issueType": {
                    "All": true,
                    "Problem": false,
                    "Action": false,
                    "Event": false,
                    "Comment": false
                },
                "priority": {
                    "All": true,
                    "High": false,
                    "Medium": false,
                    "Low": false
                },
                "targetDate": "",
                "personResponsible": "",
                "orderField": "id",
                "sortOrder": "ASC",
                "coreUserEmail": ""
            };


            $.ajax({
                type: "POST",
                url: urlProject,
                dataType: "json",
                contentType: "application/json; charset=utf-8",
                data: JSON.stringify(dataEmail),
            })
              .done(function (str) {
                  var length = Object.keys(str).length;
                  var DLId = [length]
                  if (length > 0) {
                      var matrix = [length];
                      for (var i = 0; i < length; i++) {
                          matrix[i] = [7];
                          DLId[i] = str[i].id;
                          matrix[i][0] = isNull(str[i].title);
                          matrix[i][1] = isNull(str[i].id);
                          matrix[i][2] = isNull(str[i].status);
                          matrix[i][3] = isNull(str[i].issueType);
                          matrix[i][4] = isNull(str[i].priority);
                          matrix[i][5] = formatDate(str[i].targetDate);
                          matrix[i][6] = isNull(str[i].personResponsible);


                      }
                  }
                  else {
                      var matrix = [["", "", "", "", "", "", ""]]
                  }
                  sessionStorage.setItem("DLId", DLId);
                  var dailyLog = new Office.TableData();
                  dailyLog.headers = ["Activity Title", "ID", "Status", "Type", "Priority", "Target", "Person Responsible"];
                  dailyLog.rows = matrix;
                  // Set the myTable in the document.
                  Office.context.document.setSelectedDataAsync(
                    dailyLog,
                    {
                        coercionType: Office.CoercionType.Table, cellFormat: [{ cells: Office.Table.All, format: { width: "auto fit" } }
                        ]
                    },
                    function (asyncResult) {
                        if (asyncResult.status == "failed") {
                            //showMessage("Action failed with error: " + asyncResult.error.message);
                        } else {
                            //showMessage("Check out your new table, then click next to learn another API call.");
                        }
                    }
                  );

                  Office.context.document.bindings.addFromSelectionAsync(
                       Office.BindingType.Table,
                               { id: "dailyLog" },
                       function (asyncResult) {
                           if (asyncResult.status == "failed") {
                               //showMessage("Action failed with error: " + asyncResult.error.message);
                           } else {
                               //app.showNotification('Binding done');
                           }
                       });
              });
        });

        $(document).on("click", '#IssueRegister', function () {
            var projectId = sessionStorage.getItem('projectId'); //when using the login-screen
            //app.showNotification(projectId);

            //var projectId = '22050'; //to test just this page
            var urlProject = host + '/api/IssueRegister/GetIssues';
            var dataEmail = {
                "projectId": projectId,
                "identifier": "",
                "title": "",
                "dateRaised": "",
                "raisedBy": "",
                "issueType": {
                    "All": true,
                    "RequestforChange": false,
                    "OffSpecification": false,
                    "ProblemConcern": false
                },
                "priority": {
                    "All": true,
                    "High": false,
                    "Medium": false,
                    "Low": false
                },
                "status": {
                    "All": true,
                    "New": false,
                    "Open": false,
                    "Closed": false
                },
                "orderField": "id",
                "sortOrder": "ASC"
            }
            $.ajax({
                type: "POST",
                url: urlProject,
                dataType: "json",
                contentType: "application/json; charset=utf-8",
                data: JSON.stringify(dataEmail),
            })
              .done(function (str) {
                  var length = Object.keys(str).length;
                  var IRId = [length];
                  if (length > 0) {
                      var matrix = [length];
                      for (var i = 0; i < length; i++) {
                          matrix[i] = [8];
                          IRId[i] = str[i].id;
                          matrix[i][0] = isNull(str[i].title);
                          matrix[i][1] = isNull(str[i].identifier);
                          matrix[i][2] = isNull(str[i].status);
                          matrix[i][3] = isNull(str[i].issueType);
                          matrix[i][4] = isNull(str[i].priority);
                          matrix[i][5] = formatDate(str[i].dateRaised);
                          matrix[i][6] = isNull(str[i].raisedBy);
                          matrix[i][7] = isTypeSeverity(str[i].severity);
                      }
                  }
                  else {
                      var matrix = [["", "", "", "", "", "", "",""]]
                  }
                  sessionStorage.setItem("IRId", IRId);
                  var issueRegister = new Office.TableData();
                  issueRegister.headers = ["Issue Title", "Issue ID", "Status", "Issue Type", "Priority", "Raised", "Raised By", "Severity"];
                  issueRegister.rows = matrix;
                  // Set the myTable in the document.
                  Office.context.document.setSelectedDataAsync(
                    issueRegister,
                    {
                        coercionType: Office.CoercionType.Table, cellFormat: [{ cells: Office.Table.All, format: { width: "auto fit" } }
                        ]
                    },
                    function (asyncResult) {
                        if (asyncResult.status == "failed") {
                            //showMessage("Action failed with error: " + asyncResult.error.message);
                        } else {
                            //showMessage("Check out your new table, then click next to learn another API call.");
                        }
                    }
                  );

                  Office.context.document.bindings.addFromSelectionAsync(
                       Office.BindingType.Table,
                               { id: "issueRegister" },
                       function (asyncResult) {
                           if (asyncResult.status == "failed") {
                               //showMessage("Action failed with error: " + asyncResult.error.message);
                           } else {
                               //app.showNotification('Binding done');
                           }
                       });
              });
        });

        $(document).on("click", '#LessonLog', function () {
            var projectId = sessionStorage.getItem('projectId'); //when using the login-screen
            //app.showNotification(projectId);

            //var projectId = '22050'; //to test just this page
            var urlProject = host + '/api/LessonLog/GetLessons';
            var dataEmail = {
                "projectId": projectId,
                "identifier": "",
                "title": "",
                "status": {
                    "All": true,
                    "New": false,
                    "Draft": false,
                    "Approval": false,
                    "Version": false
                },
                "lessonType": {
                    "All": true,
                    "Project": false,
                    "Corporate": false,
                    "Program": false
                },
                "priority": {
                    "All": true,
                    "High": false,
                    "Medium": false,
                    "Low": false
                },
                "dateLogged": "",
                "loggedBy": "",
                "sortField": "",
                "sortOrder": ""
            }


            $.ajax({
                type: "POST",
                url: urlProject,
                dataType: "json",
                contentType: "application/json; charset=utf-8",
                data: JSON.stringify(dataEmail),
            })
              .done(function (str) {
                  var length = Object.keys(str).length;
                  var LLId = [length];
                  if (length > 0) {
                      var matrix = [length];
                      for (var i = 0; i < length; i++) {
                          matrix[i] = [7];
                          LLId[i] = str[i].id;
                          matrix[i][0] = isNull(str[i].title);
                          matrix[i][1] = isNull(str[i].identifier);
                          matrix[i][2] = isNull(str[i].status);
                          matrix[i][3] = isNull(str[i].lessonType);
                          matrix[i][4] = isNull(str[i].priority);
                          matrix[i][5] = formatDate2(str[i].dateLogged);
                          matrix[i][6] = isNull(str[i].loggedBy);

                      }
                  }
                  else {
                      var matrix = [["", "", "", "", "", "", ""]]
                  }
                  sessionStorage.setItem("LLId", LLId);
                  var lessonLog = new Office.TableData();
                  lessonLog.headers = ["Lesson Title", "Lesson ID", "Status", "Type", "Priority", "Logged", "Logged By"];
                  lessonLog.rows = matrix;
                  // Set the myTable in the document.
                  Office.context.document.setSelectedDataAsync(
                    lessonLog,
                    {
                        coercionType: Office.CoercionType.Table, cellFormat: [{ cells: Office.Table.All, format: { width: "auto fit" } }
                        ]
                    },
                    function (asyncResult) {
                        if (asyncResult.status == "failed") {
                            //showMessage("Action failed with error: " + asyncResult.error.message);
                        } else {
                            //showMessage("Check out your new table, then click next to learn another API call.");
                        }
                    }
                  );

                  Office.context.document.bindings.addFromSelectionAsync(
                       Office.BindingType.Table,
                               { id: "lessonLog" },
                       function (asyncResult) {
                           if (asyncResult.status == "failed") {
                               //showMessage("Action failed with error: " + asyncResult.error.message);
                           } else {
                               //app.showNotification('Binding done');
                           }
                       });
              });
        });

        $(document).on("click", '#ProductDescriptions', function () {
            var projectId = sessionStorage.getItem('projectId'); //when using the login-screen

            //var projectId = '22050'; //to test just this page
            var urlProject = host + '/api/ProductDescription/GetAllProductDescription?projectId=' + projectId;

            $.ajax({
                type: 'GET',
                url: urlProject,
                dataType: "json",
                contentType: "application/json; charset=utf-8",

            })
              .done(function (str) {
                  var PdId = [str.length];
                  if (str.length > 0) {
                      var matrix = [str.length];
                      for (var i = 0; i < str.length; i++) {
                          matrix[i] = [5];
                          PdId[i] = str[i].Id;
                          matrix[i][0] = isNull(str[i].Title);
                          matrix[i][1] = isNull(str[i].Identifier);
                          matrix[i][2] = isNull(str[i].Status);
                          matrix[i][3] = isNull(str[i].ProductCategory);
                          matrix[i][4] = isNull(str[i].Version);
                      }
                  } else {
                      var matrix = [["", "", "", "", "", ""]]
                  }
                  sessionStorage.setItem("PdId", PdId);
                  var productDescriptions = new Office.TableData();
                  productDescriptions.headers = ["Title", "Identifier", "Workflow Status", "Item Type", "Version"];
                  productDescriptions.rows = matrix;

                  // Set the myTable in the document.
                  Office.context.document.setSelectedDataAsync(
                    productDescriptions,
                    {
                        coercionType: Office.CoercionType.Table, cellFormat: [{ cells: Office.Table.All, format: { width: "auto fit" } }
                        ]
                    },
                    function (asyncResult) {
                        if (asyncResult.status == "failed") {
                            //showMessage("Action failed with error: " + asyncResult.error.message);
                        } else {
                            //showMessage("Check out your new table, then click next to learn another API call.");
                        }
                    }
                  );

                  Office.context.document.bindings.addFromSelectionAsync(
                       Office.BindingType.Table,
                               { id: "productDescriptions" },
                       function (asyncResult) {
                           if (asyncResult.status == "failed") {
                               //showMessage("Action failed with error: " + asyncResult.error.message);
                           } else {
                               //app.showNotification('Binding done');
                           }
                       });
              });
        });

        $(document).on("click", '#QualityRegister', function () {
            var projectId = sessionStorage.getItem('projectId'); //when using the login-screen
            //app.showNotification(projectId);

            //var projectId = '22050'; //to test just this page
            var urlProject = host + '/api/QualityRegister/GetQualityRegister';
            var dataEmail = {
                "projectId": projectId,
                "title": "",
                "identifier": "",
                "qualityActivityPlanDate": "",
                "completionQualityActivityPlanDate": "",
                "responsibleName": "",
                "sortField": "id",
                "sortOrder": "ASC"
            }


            $.ajax({
                type: "POST",
                url: urlProject,
                dataType: "json",
                contentType: "application/json; charset=utf-8",
                data: JSON.stringify(dataEmail),
            })
              .done(function (str) {
                  var length = Object.keys(str).length;
                  var QRId = [length];
                  if (length > 0) {
                      var matrix = [length];
                      for (var i = 0; i < length; i++) {
                          matrix[i] = [5];
                          QRId[i] = str[i].id;
                          matrix[i][0] = isNull(str[i].title);
                          matrix[i][1] = isNull(str[i].id);
                          matrix[i][2] = isNull(str[i].qualityActivityPlanDate);
                          matrix[i][3] = formatDate(str[i].completionQualityActivityPlanDate);
                          matrix[i][4] = isNull(str[i].responsibleName);


                      }
                  }
                  else {
                      var matrix = [["", "", "", "", ""]]
                  }
                  sessionStorage.setItem("QRId", QRId);
                  var qualityRegister = new Office.TableData();
                  qualityRegister.headers = ["Title", "ID", "Quality Activity Date", "Completion Date", "Responsible Name"];
                  qualityRegister.rows = matrix;
                  // Set the myTable in the document.
                  Office.context.document.setSelectedDataAsync(
                    qualityRegister,
                    {
                        coercionType: Office.CoercionType.Table, cellFormat: [{ cells: Office.Table.All, format: { width: "auto fit" } }
                        ]
                    },
                    function (asyncResult) {
                        if (asyncResult.status == "failed") {
                            //showMessage("Action failed with error: " + asyncResult.error.message);
                        } else {
                            //showMessage("Check out your new table, then click next to learn another API call.");
                        }
                    }
                  );

                  Office.context.document.bindings.addFromSelectionAsync(
                       Office.BindingType.Table,
                               { id: "qualityRegister" },
                       function (asyncResult) {
                           if (asyncResult.status == "failed") {
                               //showMessage("Action failed with error: " + asyncResult.error.message);
                           } else {
                               //app.showNotification('Binding done');
                           }
                       });
              });
        });

        $(document).on("click", '#Reports', function () {
            var projectId = sessionStorage.getItem('projectId'); //when using the login-screen
            //app.showNotification(projectId);

            //var projectId = '22050'; //to test just this page
            var urlProject = host + '/api/ReportCard/GetReportRegister';
            var dataEmail = {
                "projectId": projectId,
                "workFlowStatus": {
                    "All": "true",
                    "New": "false",
                    "Draft": "false",
                    "Approval": "false",
                    "Version": "false"
                },
                "sortField": "title",
                "sortOrder": "ASC",
                "title": "",
                "identifier": "",
                "version": "",
                "date": ""
            }
            $.ajax({
                type: "POST",
                url: urlProject,
                dataType: "json",
                contentType: "application/json; charset=utf-8",
                data: JSON.stringify(dataEmail),
            })
              .done(function (str) {
                  var length = Object.keys(str).length;
                  var RId = [lenght];
                  if (length > 0) {
                      var matrix = [length];
                      for (var i = 0; i < length; i++) {
                          matrix[i] = [5];
                          RId[i] = str[i].id;
                          matrix[i][0] = isNull(str[i].title);
                          matrix[i][1] = isNull(str[i].identifier);
                          matrix[i][2] = isNull(str[i].workFlowStatus);
                          matrix[i][3] = formatDate(str[i].date);
                          matrix[i][4] = isNull(str[i].version);
                      }
                  }
                  else {
                      var matrix = [["", "", "", "", ""]]
                  }
                  sessionStorage.setItem("RId", RId);
                  var reports = new Office.TableData();
                  reports.headers = ["Title", "Identifier", "Workflow Status", "Date", "Version"];
                  reports.rows = matrix;
                  // Set the myTable in the document.
                  Office.context.document.setSelectedDataAsync(
                    reports,
                    {
                        coercionType: Office.CoercionType.Table, cellFormat: [{ cells: Office.Table.All, format: { width: "auto fit" } }
                        ]
                    },
                    function (asyncResult) {
                        if (asyncResult.status == "failed") {
                            //showMessage("Action failed with error: " + asyncResult.error.message);
                        } else {
                            //showMessage("Check out your new table, then click next to learn another API call.");
                        }
                    }
                  );

                  Office.context.document.bindings.addFromSelectionAsync(
                       Office.BindingType.Table,
                               { id: "reports" },
                       function (asyncResult) {
                           if (asyncResult.status == "failed") {
                               //showMessage("Action failed with error: " + asyncResult.error.message);
                           } else {
                               //app.showNotification('Binding done');
                           }
                       });
              });
        }); //to do

        $(document).on("click", '#RiskRegister', function () {
            var projectId = sessionStorage.getItem('projectId');
            var urlProject = host + '/api/RiskRegister/GetRiskRegister';
            var dataEmail = {
                "projectId": projectId,
                "identifier": "",
                "title": "",
                "riskStatus": {
                    "All": true,
                    "New": false,
                    "Active": false,
                    "Closed": false
                },
                "riskType": {
                    "All": true,
                    "Threat": false,
                    "Opportunity": false
                },
                "dateRegistered": "",
                "riskOwner": "",
                "sortField": "id",
                "sortOrder": "ASC"
            }

            $.ajax({
                type: "POST",
                url: urlProject,
                dataType: "json",
                contentType: "application/json; charset=utf-8",
                data: JSON.stringify(dataEmail),
            })
              .done(function (str) {
                  var length = Object.keys(str).length;
                  var RRId = [length];
                  if (length > 0) {
                      var matrix = [length];
                      for (var i = 0; i < length; i++) {
                          matrix[i] = [6];
                          RRId[i] = str[i].id;
                          matrix[i][0] = isNull(str[i].title);
                          matrix[i][1] = isNull(str[i].identifier);
                          matrix[i][2] = isNull(str[i].riskStatus);
                          matrix[i][3] = isNull(str[i].riskType);
                          matrix[i][4] = formatDate(str[i].dateRegistered);
                          matrix[i][5] = isNull(str[i].riskOwner);
                      }
                  }
                  else {
                      var matrix = [["", "", "", "", "", ""]]
                  }
                  sessionStorage.setItem("RRId", RRId);
                  var riskRegister = new Office.TableData();
                  riskRegister.headers = ["Risk Title", "Risk ID", "Status", "Risk Type", "Date", "Risk Owner"];
                  riskRegister.rows = matrix;
                  // Set the myTable in the document.
                  Office.context.document.setSelectedDataAsync(
                    riskRegister,
                    {
                        coercionType: Office.CoercionType.Table, cellFormat: [{ cells: Office.Table.All, format: { width: "auto fit" } }, { cells: { column: 4 }, format: { numberFormat: "dd-mm-yyyy" } }
                        ]
                    },
                    function (asyncResult) {
                        if (asyncResult.status == "failed") {
                            //showMessage("Action failed with error: " + asyncResult.error.message);
                        } else {
                            //showMessage("Check out your new table, then click next to learn another API call.");
                        }
                    }
                  );

                  Office.context.document.bindings.addFromSelectionAsync(
                       Office.BindingType.Table,
                               { id: "riskRegister" },
                       function (asyncResult) {
                           if (asyncResult.status == "failed") {
                               //showMessage("Action failed with error: " + asyncResult.error.message);
                           } else {
                               //app.showNotification('Binding done');
                           }
                       });
              });
        });

        $(document).on("click", "#Refresh", function () {
            var table = new Office.TableData();
            //table.headers = ["a", "b", "a", "b", "a", "b"]
            table.rows = [["Seattle", "WA", "test1", "test2", "test3", "test4"], ["Seattle", "WA", "test1", "test2", "test3", "test4"]];
            //var table = [["Seattle", "WA", "test1", "test2", "test3", "test4"], ["Seattle", "WA", "test1", "test2", "test3", "test4"]];

            //updateCmdb();
            Office.context.document.bindings.getByIdAsync("dailyLog", function (asyncResult) {
                var binding = asyncResult.value;
                binding.deleteAllDataValuesAsync();
                binding.setDataAsync(table, { coercionType: "table" });
                binding.setDataAsync(table, { coercionType: "table", startRow: 1 });
                binding.setDataAsync(table, { coercionType: "table", startRow: 2 });
                binding.setDataAsync(table, { coercionType: "table", startRow: 3 });
            });

        });

        $(document).on("click", "#Publish", function () {
            //publishCmdb();
            //publishDailyLog();
            publishRiskRegister();
            publishIssueRegister();
            publishLessonLog();
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
        if (date == null) { return '' }
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
    }

    //Risk Register
    function isRiskStatus(status) {
        if (status == 'New') return '0';
        else if (status == 'Active') return '1';
        else if (status == 'Closed') return '2';
        else {
            app.showNotification('Wrong Risk Status. Accepted values are "New", "Active", "Closed".');
            return null
        }
    }

    function isRiskType(type) {
        if (type == 'Threat') return 0;
        else if (type == 'Opportunity') return 1;
        else {
            app.showNotification('Wrong Risk Type. Accepted values are "Threat", "Opportunity".')
        }
    }

    //Issue Register
    function isIssueStatus(status) {
        if (status == "New") return "0";
        else if (status == "Open") return "1";
        else if (status == "Closed") return "2";
        else app.showNotification('Wrong Issue Status. Accepted values are "New", "Open" and "Closed".')
    };

    function isIssueType(type) {
        if (type == "Request For Change") return "0";
        else if (type == "Off Specification") return "1";
        else if (type == "Problem Concern") return "2";
        else app.showNotification('Wrong Issue Type. Accepted values are "Request For Change", "Off Specification" and "Problem Concern".');
    }

    function isPriority(priority) {
        if (priority == "Low") return "2";
        else if (priority == "Medium") return "1";
        else if (priority == "High") return "0";
        else app.showNotification('Wrong Priority. Accepted values are "Low", "Medium" and "High".');
    }

    function isSeverity(severity) {
        if (severity == "Team Manager") return "0";
        else if (severity == "Project Manager") return "1";
        else if (severity == "Project Board") return "2";
        else if (severity == "Corporate / Program Management") return "3";
        else if (severity == "" || severity == null) return "";
        else app.showNotification('Wrong Severity. Accepted values are "Team Manager", "Project Manager", "Project Board" and "Corporate / Program Management".')
    }

    //Lesson Log
    function isLessonType(type) {
        if (type == "Project") return "0";
        else if (type == "Program") return "1";
        else if (type == "Corporate") return "2";
        else app.showNotification('Wrong Lesson Type. Accepted values are "Project", "Program" and "Corporate".')
    }

    function isLessonStatus(status) {
        if (status == "New") return "0";
        else if (status == "Draft") return "1";
        else if (status == "Approval") return "2";
        else if (status == "Version") return "3";
        else app.showNotification('Wrong Lesson Status. Accepted values are "New", "Draft", "Approval" and "Version".')
    }

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

    function updateCmdb() {
        var projectId = sessionStorage.getItem('projectId');
        var urlProject = host + '/api/ConfigManagerDb/GetConfigurationManagerRegister';
        var dataEmail = {
            "projectId": projectId,
            "sortField": "title",
            "sortOrder": "DESC",
            "status": {
                "All": "true",
                "PendingDevelopment": "false",
                "InDevelopment": "false",
                "InReview": "false",
                "Approved": "false",
                "HandedOver": "false"
            },
            "type": {
                "All": "true",
                "Component": "false",
                "Product": "false",
                "Release": "true"
            },
            "title": "",
            "identifier": ""
        }


        $.ajax({
            type: "POST",
            url: host + "/api/ConfigManagerDb/GetConfigurationManagerRegister",
            dataType: "json",
            contentType: "application/json; charset=utf-8",
            data: JSON.stringify(dataEmail),
        })
          .done(function (str) {
              var length = Object.keys(str).length;
              if (length > 0) {
                  var matrix = [length];
                  for (var i = 0; i < length; i++) {
                      matrix[i] = [6];
                      matrix[i][0] = isNull(str[i].id);
                      matrix[i][1] = isNull(str[i].title);
                      matrix[i][2] = isNull(str[i].producer);
                      matrix[i][3] = isNull(str[i].location);
                      matrix[i][4] = isNull(str[i].status);
                      matrix[i][5] = isNull(str[i].type);
                  }
              }
              else {
                  var matrix = [["", "", "", "", "", ""]]
              }
              var table = new Office.TableData();
              table.rows = matrix;

              Office.context.document.bindings.getByIdAsync("cmdb", function (asyncResult) {
                  var binding = asyncResult.value;
                  binding.deleteAllDataValuesAsync();
                  binding.setDataAsync(table, { coercionType: Office.CoercionType.Table });
              });
          })
    };

    function publishCmdb() {
        Office.select("bindings#cmdb").getDataAsync({ coercionType: 'table' }, function (result) {
            var binding = result.value.rows;
            var CmdbId = sessionStorage.getItem('CmdbId');
            var projectId = sessionStorage.getItem('projectId');
            var urlProject = host + '/api/ConfigManagerDb/PostConfigManDb/';
            for (var i = 0; i < CmdbId.length; i++) {
                var dataEmail = {
                    "id": null,
                    "projectId": "5",
                    "identifier": binding[i][1],
                    "title": binding[i][0],
                    "status": binding[i][2],//
                    "type": binding[i][3],//
                    "stage": "The Stage"
                };
                $.ajax({
                    type: "POST",
                    url: urlProject,
                    dataType: "json",
                    contentType: "application/json; charset=utf-8",
                    data: JSON.stringify(dataEmail),
                })
            }
        })
    }

    //ok
    function publishRiskRegister() {
        Office.select("bindings#riskRegister").getDataAsync({ coercionType: 'table' }, function (result) {
            var binding = result.value.rows;
            var RRId = sessionStorage.getItem('RRId');
            var projectId = sessionStorage.getItem('projectId');
            var urlProject = host + '/api/RiskRegister/PostRiskRegisterHeader';
            for (var i = 0; i < RRId.length; i++) {
                //app.showNotification(typeof projectId);
                var dataEmail = {
                    "id": binding[i][1],
                    "projectId": projectId,
                    "title": binding[i][0],
                    "riskStatus": isRiskStatus(binding[i][2]),
                    "riskType": isRiskType(binding[i][3]),
                    "riskCategory": "33",
                    "proximity": "15",
                    "author": "Kurt",
                    "riskOwner": binding[i][5],
                    "dateRegistered": convertDate(binding[i][4]),
                    "version": "1.1",
                    "workflowStatus": "2"
                };
                $.ajax({
                    type: "POST",
                    url: urlProject,
                    dataType: "json",
                    contentType: "application/json; charset=utf-8",
                    data: JSON.stringify(dataEmail),
                })
            }

        })
    };

    //ok
    function publishIssueRegister() {
        Office.select("bindings#issueRegister").getDataAsync({ coercionType: 'table' }, function (result) {
            var binding = result.value.rows;
            var IRId = sessionStorage.getItem('IRId');
            var projectId = sessionStorage.getItem('projectId');
            var urlProject = host + '/api/IssueRegister/PostIssueHeader';
            for (var i = 0; i < IRId.length; i++) {
                var dataEmail = {
                    "id": binding[i][1],
                    "projectId": projectId,
                    "title": binding[i][0],
                    "status": isIssueStatus(binding[i][2]),
                    "issueType": isIssueType(binding[i][3]),
                    "priority": isPriority(binding[i][4]),
                    "severity": isSeverity(binding[i][7]),
                    "dateRaised": convertDate(binding[i][5]),
                    "raisedBy": binding[i][6]
                }
                $.ajax({
                    type: "POST",
                    url: urlProject,
                    dataType: "json",
                    contentType: "application/json; charset=utf-8",
                    data: JSON.stringify(dataEmail),
                })
            }

        })
    };

    //ok
    function publishLessonLog() {
        Office.select("bindings#lessonLog").getDataAsync({ coercionType: 'table' }, function (result) {
            var binding = result.value.rows;
            var LLId = sessionStorage.getItem('LLId');
            var projectId = sessionStorage.getItem('projectId');
            var urlProject = host + '/api/LessonLog/PostLessonLogHeader';
            for (var i = 0; i < LLId.length; i++) {
                var dataEmail = {
                    "id": binding[i][1],
                    "projectId": projectId,
                    "title": binding[i][0],
                    "statusId": isLessonStatus(binding[i][2]),
                    "lessonTypeId": isLessonType(binding[i][3]),
                    "priorityId": isPriority(binding[i][4]),
                    "version": "",
                    "dateLogged": convertDate2(binding[i][5]),
                    "loggedBy": binding[i][6]
                };
                $.ajax({
                    type: "POST",
                    url: urlProject,
                    dataType: "json",
                    contentType: "application/json; charset=utf-8",
                    data: JSON.stringify(dataEmail),
                })
            }

        })
    }

    function publishQualityRegister() {
        Office.select("bindings#qualityRegister").getDataAsync({ coercionType: 'table' }, function (result) {
            var binding = result.value.rows;
            var QRId = sessionStorage.getItem('QRId');
            var projectId = sessionStorage.getItem('projectId');
            var urlProject = host + '/api/LessonLog/PostLessonLogHeader';
            for (var i = 0; i < QRId.length; i++) {
                var dataEmail = {
                    "id": binding[i][1],
                    "projectId": projectId,
                    "title": binding[i][0],
                    "statusId": isLessonStatus(binding[i][2]),
                    "lessonTypeId": isLessonType(binding[i][3]),
                    "priorityId": isPriority(binding[i][4]),
                    "version": "",
                    "dateLogged": convertDate2(binding[i][5]),
                    "loggedBy": binding[i][6]
                };
                $.ajax({
                    type: "POST",
                    url: urlProject,
                    dataType: "json",
                    contentType: "application/json; charset=utf-8",
                    data: JSON.stringify(dataEmail),
                })
            }

        })
    }

    return app;
})();