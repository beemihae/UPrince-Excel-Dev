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

