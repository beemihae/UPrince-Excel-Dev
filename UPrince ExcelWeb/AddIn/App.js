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

        $(document).on("click", '#Cmdb', function () {
            //id excluded
            var projectId = localStorage.getItem('projectId'); //when using the login-screen
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
                  localStorage.setItem("CmdbId", CmdbId);
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
            activateWorksheet("DailyLog");
            dailyLogGET();
        });

        $(document).on("click", '#IssueRegister', function () {
            var projectId = localStorage.getItem('projectId'); //when using the login-screen
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
                      var matrix = [["", "", "", "", "", "", "", ""]]
                  }
                  localStorage.setItem("IRId", IRId);
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
            var projectId = localStorage.getItem('projectId'); //when using the login-screen
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
                  localStorage.setItem("LLId", LLId);
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
            activateWorksheet("ProductDescription");
            productDescriptionGET();

        });

        $(document).on("click", '#QualityRegister', function () {
            var projectId = localStorage.getItem('projectId'); //when using the login-screen
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
                  localStorage.setItem("QRId", QRId);
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
            var projectId = localStorage.getItem('projectId'); //when using the login-screen
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
                  localStorage.setItem("RId", RId);
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
            activateWorksheet("RiskRegister");
            riskRegisterGET();
        });

        $(document).on("click", "#Refresh", function () {
            //deleteTable('ProductDescription');
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
    };

    function formatDate3(date) {
        if (date == null) return '';
        else return new Date((date - (25567 + 2)) * 86400 * 1000)
    };

    //Risk Register
    function isRiskStatus(status) {
        if (status == 'New') return '0';
        else if (status == 'Active') return '1';
        else if (status == 'Closed') return '2';
        else {
            //app.showNotification('Wrong Risk Status. Accepted values are "New", "Active", "Closed".');
            return null
        }
    }

    function isRiskType(type) {
        if (type == 'Threat') return 0;
        else if (type == 'Opportunity') return 1;
        else {
            //app.showNotification('Wrong Risk Type. Accepted values are "Threat", "Opportunity".')
            return null;
        }
    }

    //Issue Register
    function isIssueStatus(status) {
        if (status == "New") return "0";
        else if (status == "Open") return "1";
        else if (status == "Closed") return "2";
        else return null;//app.showNotification('Wrong Issue Status. Accepted values are "New", "Open" and "Closed".')
    };

    function isIssueType(type) {
        if (type == "Request For Change") return "0";
        else if (type == "Off Specification") return "1";
        else if (type == "Problem Concern") return "2";
        else return null //app.showNotification('Wrong Issue Type. Accepted values are "Request For Change", "Off Specification" and "Problem Concern".');
    }

    function isPriority(priority) {
        if (priority == "Low") return "2";
        else if (priority == "Medium") return "1";
        else if (priority == "High") return "0";
        else return null //app.showNotification('Wrong Priority. Accepted values are "Low", "Medium" and "High".');
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
        else return null //app.showNotification('Wrong Lesson Type. Accepted values are "Project", "Program" and "Corporate".')
    }

    function isLessonStatus(status) {
        if (status == "New") return "0";
        else if (status == "Draft") return "1";
        else if (status == "Approval") return "2";
        else if (status == "Version") return "3";
        else return null //app.showNotification('Wrong Lesson Status. Accepted values are "New", "Draft", "Approval" and "Version".')
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
        var projectId = localStorage.getItem('projectId');
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
            var CmdbId = localStorage.getItem('CmdbId');
            var projectId = localStorage.getItem('projectId');
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

    function publishIssueRegister() {
        Office.select("bindings#issueRegister").getDataAsync({ coercionType: 'table' }, function (result) {
            var binding = result.value.rows;
            var IRId = localStorage.getItem('IRId');
            var projectId = localStorage.getItem('projectId');
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


    function publishLessonLog() {
        Office.select("bindings#lessonLog").getDataAsync({ coercionType: 'table' }, function (result) {
            var binding = result.value.rows;
            var LLId = localStorage.getItem('LLId');
            var projectId = localStorage.getItem('projectId');
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
            var QRId = localStorage.getItem('QRId');
            var projectId = localStorage.getItem('projectId');
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
                //app.showNotification(rangeAddress);
                localStorage.setItem('rangeAddress', rangeAddress);
                var sheetName = range.substring(0, range.indexOf('!'));
                localStorage.setItem('sheetName', sheetName);

                //tableDataRange.address.delete();
            });
        }).catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });

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
        /*function deleteTable(tableName) {
            Excel.run(function (ctx) {
                var table = ctx.workbook.tables.getItem(tableName);
                var tableRange = table.getRange();
                tableRange.delete();
            }).catch(function (error) {
                console.log("Error: " + error);
                if (error instanceof OfficeExtension.Error) {
                    console.log("Debug info: " + JSON.stringify(error.debugInfo));
                }
            });
        }*/
    };

    //risk register
    function riskRegisterGET() {
        //deleteTable('riskRegister');
        var projectId = localStorage.getItem('projectId');
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
                      matrix[i] = [12];
                      RRId[i] = str[i].id;
                      matrix[i][0] = isNull(str[i].title);
                      matrix[i][1] = isNull(str[i].identifier);
                      matrix[i][2] = isNull(str[i].riskStatus);
                      matrix[i][3] = isNull(str[i].riskType);
                      matrix[i][4] = formatDate(str[i].dateRegistered);
                      matrix[i][5] = isNull(str[i].riskOwner);
                      var urlProject = host + '/api/RiskRegister/GetRiskRegisterEntry?riskId=' + str[i].id + '&projectId=' + projectId;
                      $.ajax({
                          type: 'GET',
                          url: urlProject,
                          async: false,
                          dataType: "json",
                          contentType: "application/json; charset=utf-8",
                      })
                       .done(function (anw) {
                           //app.showNotification(anw.selectedImpactInherent);
                           matrix[i][6] = isNull(anw.selectedImpactInherent);
                           //app.showNotification(matrix[i][6]);
                           matrix[i][7] = isNull(anw.selectedImpactResidual);
                           matrix[i][8] = isNull(anw.selectedProbabilityInherent);
                           matrix[i][9] = isNull(anw.selectedProbabilityResidual);
                           matrix[i][10] = isNull(anw.expectedValueInherent);
                           matrix[i][11] = isNull(anw.expectedValueResidual);
                       })
                  }
              }
              else {
                  var matrix = [["", "", "", "", "", "", "", "", "", "", "", ""]]
              };
              getRiskRegister(projectId, RRId[0]);
              localStorage.setItem("RRId", RRId);

              Excel.run(function (ctx) {
                  var tables = ctx.workbook.tables;
                  var tableRows = tables.getItem('riskRegister').rows
                  for (var i = 0; i < matrix.length; i++) {
                      var line = [1];
                      line[0] = matrix[i];
                      tableRows.add(null, line);
                  };
                  return ctx.sync().then(function () {
                      showMessage("Success! My monthly expense table created! Select the arrow button to see how to remove the table.");
                  })
                   .catch(function (error) {
                       showMessage(JSON.stringify(error));
                   });
              });
          });
    };

    function getRiskRegister(projectId, riskId) {
        var urlProject = host + '/api/RiskRegister/GetRiskRegisterEntry?riskId=' + riskId + '&projectId=' + projectId;
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
                 ctx.workbook.worksheets.getItem('Values').getRange("C1:C" + Object.keys(str.impact).length).values = riskValuesImpact(str)/*[[1], [2], [1], [2], [1]] //str.impact[0].State*/;
                 ctx.workbook.worksheets.getItem('Values').getRange("D1:D" + Object.keys(str.probability).length).values = riskValuesProb(str)/*[[1], [2], [1], [2], [1]] //str.impact[0].State*/;
                 ctx.workbook.worksheets.getItem('Values').getRange("A1:A3").values = [["New"], ["Active"], ["Closed"]]/*[[1], [2], [1], [2], [1]] //str.impact[0].State*/;
                 ctx.workbook.worksheets.getItem('Values').getRange("B1:B2").values = [["Threat"], ["Opportunity"]]/*[[1], [2], [1], [2], [1]] //str.impact[0].State*/;
                 return ctx.sync().then(function () {
                     //console.log("Success! Insert range in A1:C3.");
                 });;
             }).catch(function (error) {
                 console.log(error);
             });
         })
    };

    function riskValuesImpact(str) {
        var val = [Object.keys(str.impact).length];
        for (var i = 0; i < Object.keys(str.impact).length; i++) {
            val[i] = [1];
            val[i][0] = str.impact[i].State;
            localStorage.setItem('riskValuesImpact' + str.impact[i].State, "" + str.impact[i].StateId);
            //val[i] = str.impact[i].State;
        }
        //app.showNotification(val[2][0]);
        return val;
    };

    function riskValuesProb(str) {
        var val = [Object.keys(str.impact).length];
        for (var i = 0; i < Object.keys(str.impact).length; i++) {
            val[i] = [1];
            val[i][0] = str.probability[i].State;
            localStorage.setItem('riskValuesProb' + str.probability[i].State, "" + str.probability[i].StateId);

            //val[i] = str.impact[i].State;
        }
        //app.showNotification(val[2][0]);
        return val;
    };

    function publishRiskRegister() {
        Excel.run(function (ctx) {
            var rows = ctx.workbook.tables.getItem("riskRegister").rows.load("values");
            return ctx.sync()
                .then(function () {
                    var RRId = localStorage.getItem('RRId');
                    var projectId = localStorage.getItem('projectId');
                    var urlProject = host + '/api/RiskRegister/PostRiskRegisterHeader';
                    var urlProject2 = host + '/api/RiskRegister/PostRiskRegisterImpact';
                    for (var i = 0; i < rows.items.length; i++) {
                        //app.showNotification(rows.items[1].values[0][1]);
                        var dataEmail = {
                            "id": rows.items[i].values[0][1],
                            "projectId": projectId,
                            "title": rows.items[i].values[0][0],
                            "riskStatus": isRiskStatus(rows.items[i].values[0][2]),
                            "riskType": isRiskType(rows.items[i].values[0][3]),
                            "riskCategory": "33",
                            "proximity": "15",
                            "author": "Kurt",
                            "riskOwner": rows.items[i].values[0][5],
                            "dateRegistered": convertDate(rows.items[i].values[0][4]),
                            "version": "1.1",
                            "workflowStatus": "2"
                        };
                        $.ajax({
                            type: "POST",
                            url: urlProject,
                            dataType: "json",
                            contentType: "application/json; charset=utf-8",
                            data: JSON.stringify(dataEmail),
                        });

                        var dataEmail2 = {
                            "riskEntryId": rows.items[i].values[0][1],
                            "impactInherent": localStorage.getItem('riskValuesImpact' + rows.items[i].values[0][6]),
                            "impactResidual": localStorage.getItem('riskValuesImpact' + rows.items[i].values[0][7]),
                            "probabilityInherent": localStorage.getItem('riskValuesProb' + rows.items[i].values[0][8]),
                            "probabilityResidual": localStorage.getItem('riskValuesProb' + rows.items[i].values[0][9]),
                            "expectedInherent": "",//isNull(rows.items[i].values[0][10]),
                            "expectedResidual": ""//isNull(rows.items[i].values[0][11])
                        };
                        $.ajax({
                            type: "POST",
                            url: urlProject2,
                            dataType: "json",
                            contentType: "application/json; charset=utf-8",
                            data: JSON.stringify(dataEmail2),
                        });
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

    //Product Description
    function productDescriptionGET() {
        //deleteTable('ProductDescription');
        var projectId = localStorage.getItem('projectId');
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
                     matrix[i] = [6];
                     PdId[i] = str[i].Id;
                     matrix[i][0] = isNull(str[i].Title);
                     matrix[i][1] = isNull(str[i].Identifier);
                     matrix[i][2] = isNull(str[i].ProductCategory);
                     matrix[i][3] = isNull(str[i].ToleranceStatus);
                     matrix[i][4] = isNull(str[i].Status);
                     matrix[i][5] = isNull(str[i].Version);
                     //matrix[i][6] = isNull(str[i].Version);
                     localStorage.setItem("ParentId" + str[i].Id, str[i].ParentId);
                 }
             } else {
                 var matrix = [["", "", "", "", "", "", ""]]
             }
             Excel.run(function (ctx) {
                 ctx.workbook.worksheets.getItem('Values').getRange("E1:E2").values = [["Internal Product"], ["External Product"]]/*[[1], [2], [1], [2], [1]] //str.impact[0].State*/;
                 ctx.workbook.worksheets.getItem('Values').getRange("F1:F4").values = [["New"], ["Draft"], ["Approval"], ["Version"]]/*[[1], [2], [1], [2], [1]] //str.impact[0].State*/;
                 return ctx.sync().then(function () {
                     //console.log("Success! Insert range in A1:C3.");
                 });;
             }).catch(function (error) {
                 console.log(error);
             });
             localStorage.setItem("PdId", PdId);

             Excel.run(function (ctx) {
                 //var productDescription = ctx.workbook.tables.add('ProductDescription!A1:F1', true);
                 //productDescription.name = 'ProductDescription';
                 //productDescription.getHeaderRowRange().values = [["Title", "Identifier", "Item Type", "Tolerance Status", "Workflow Status", "Version"]];
                 //var tableRows = productDescription.rows;
                 var tables = ctx.workbook.tables;
                 var tableRows = tables.getItem('ProductDescription').rows
                 for (var i = 0; i < matrix.length; i++) {
                     var line = [1];
                     line[0] = matrix[i];
                     tableRows.add(null, line);
                 };
                 return ctx.sync().then(function () {
                     showMessage("Success! My monthly expense table created! Select the arrow button to see how to remove the table.");
                 })
                  .catch(function (error) {
                      showMessage(JSON.stringify(error));
                  });
             });
         });

    };

    function publishProductDescription() {
        Excel.run(function (ctx) {
            var rows = ctx.workbook.tables.getItem("ProductDescription").rows.load("values");
            return ctx.sync()
                .then(function () {
                    var PdId = localStorage.getItem('PdId');
                    var projectId = localStorage.getItem('projectId');
                    var urlProject = host + '/api/productdescription/PostProductDescription';
                    for (var i = 0; i < rows.items.length; i++) {
                        //app.showNotification(rows.items.length);
                        var dataEmail = {
                            "id": rows.items[i].values[0][1],
                            "title": rows.items[i].values[0][0],
                            "productcategory": isProductCategory(rows.items[i].values[0][2]),
                            "version": rows.items[i].values[0][5],
                            "status": isWorkflowStatus(rows.items[i].values[0][4]),
                            "tolerancestatus": isToleranceStatus(rows.items[i].values[0][3]),
                            "parentid": localStorage.getItem("ParentId" + rows.items[i].values[0][1]),
                            "projectid": projectId
                        };
                        $.ajax({
                            type: "POST",
                            url: urlProject,
                            dataType: "json",
                            contentType: "application/json; charset=utf-8",
                            data: JSON.stringify(dataEmail),
                        });

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

    function isProductCategory(category) {
        if (category == "External Product") return "1";
        else return "0";
    };

    function isToleranceStatus(tolerance) {
        if (tolerance == "Within Tolerance") return "0";
        else if (tolerance == "Tolerance Limit") return "1";
        else return "2";
    };

    function isWorkflowStatus(status) {
        if (status == "New") return "0";
        else if (status == "Draft") return "1";
        else if (status == "Approval") return "2";
        else return "3";
    };

    //Daily Log
    function dailyLogGET() {

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
                    })
                     .catch(function (error) {
                         showMessage(JSON.stringify(error));
                     });
                });
                getDailyLog(projectId, DlId[0]);

            });
    };

    function getDailyLog(projectId, dailyLogId) {
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
                 ctx.workbook.worksheets.getItem('Values').getRange("G1:G" + Object.keys(str.contextList).length).values = dailyLogContext(str)/*[[1], [2], [1], [2], [1]] //str.impact[0].State*/;
                 ctx.workbook.worksheets.getItem('Values').getRange("I1:I" + Object.keys(str.personnelContacts).length).values = dailyLogUsers(str)/*[[1], [2], [1], [2], [1]] //str.impact[0].State*/;
                 ctx.workbook.worksheets.getItem('Values').getRange("H1:H6").values = [["Inbox"], ["Next"], ["Waiting"], ["Schedule"], ["Someday"], ["Done"]]/*[[1], [2], [1], [2], [1]] //str.impact[0].State*/;
                 ctx.workbook.worksheets.getItem('Values').getRange("J1:J6").values = [["Problem"], ["Action"], ["Event"], ["Comment"], ["Decision"], ["Reference"]]/*[[1], [2], [1], [2], [1]] //str.impact[0].State*/;
                 ctx.workbook.worksheets.getItem('Values').getRange("K1:K7").values = [["5 min"], ["15 min"], ["30 min"], ["1 hr"], ["2 hr"], ["4 hr"], ["8 hr"]]/*[[1], [2], [1], [2], [1]] //str.impact[0].State*/;
                 ctx.workbook.worksheets.getItem('Values').getRange("L1:L5").values = [["Mild"], ["Reasonable"], ["Demanding"], ["Very Demanding"], ["Extreme"]]/*[[1], [2], [1], [2], [1]] //str.impact[0].State*/;
                 ctx.workbook.worksheets.getItem('Values').getRange("M1:M" + Object.keys(str.project).length).values = dailyLogProject(str);
                 return ctx.sync().then(function () {
                     //console.log("Success! Insert range in A1:C3.");
                 });;
             }).catch(function (error) {
                 console.log(error);
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
            localStorage.setItem('dailyLogUsers' + str.personnelContacts[i].State, "" + str.personnelContacts[i].Index);
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
                    var urlProject = host + '/api/DailyLog/PostDailyLogHeader';
                    var urlProject2 = host + '/api/DailyLog/PostDailyLogInvolvedTiming';
                    for (var i = 0; i < rows.items.length; i++) {
                        //app.showNotification(rows.items[1].values[0][1]);
                        var dataEmail = {
                            "id": rows.items[i].values[0][2],
                            "projectId": localStorage.getItem("dailyLogProject" + rows.items[i].values[0][0]),
                            "activityTypeId": isResponseStatus(rows.items[i].values[0][7]),
                            "title": isNull(rows.items[i].values[0][1]),
                            "order": "0",
                            "coreUserEmail": localStorage.getItem("email"),
                            "authorEmail": localStorage.getItem("email"),
                            "context": localStorage.getItem("dailyLogContext" + rows.items[i].values[0][3]),
                            "energy": isEnergy(rows.items[i].values[0][9])
                        };
                        $.ajax({
                            type: "POST",
                            url: urlProject,
                            dataType: "json",
                            contentType: "application/json; charset=utf-8",
                            data: JSON.stringify(dataEmail),
                        });
                        var dataEmail2 = {
                            "dailyLogId": rows.items[i].values[0][1],
                            "responsible": localStorage.getItem('dailyLogUsers' + rows.items[i].values[0][6]),
                            "responseStatus": isResponseStatus(rows.items[i].values[0][7]),
                            "targetDate": formatDate3(rows.items[i].values[0][4]),
                            "time": isTime(rows.items[i].values[0][10])
                        };
                        $.ajax({
                            type: "POST",
                            url: urlProject2,
                            dataType: "json",
                            contentType: "application/json; charset=utf-8",
                            data: JSON.stringify(dataEmail2),
                        });
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

    function isEnergy(type) {
        if (type == "Mild") return "0";
        else if (type == "Reasonable") return "1";
        else if (type == "Demanding") return "2";
        else if (type == "Very Demanding") return "3";
        else if (type == "Extreme") return "4";
        else return null;
    };

    function isTime(type) {
        if (type = "5 min") return "0";
        else if (type = "15 min") return "1";
        else if (type = "30 min") return "2";
        else if (type = "1 hr") return "3";
        else if (type = "2 hr") return "4";
        else if (type = "4 hr") return "5";
        else if (type = "8 hr") return "6";
        else return null;
    };

    function isResponseStatus(type) {
        if (type = "Problem") return "0";
        else if (type = "Action") return "1";
        else if (type = "Event") return "2";
        else if (type = "Comment") return "3";
        else if (type = "Decision") return "4";
        else if (type = "Reference") return "5";
        else return null;
    };
    return app;
})();