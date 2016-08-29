(function () {

    Office.initialize = function (reason) {
        app.initialize();
        var host = 'https://uprinceuatapi.azurewebsites.net'
        var userEmail = localStorage.getItem('email');

        //If you need to initialize something you can do so here.
        $(document).on("click", "#addProject", function () {
            var dataEmail = {
                "id": null,
                "title": document.getElementById("project").value,
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
                dataType: "json",
                contentType: "application/json; charset=utf-8",
                data: JSON.stringify(dataEmail),
            })
                .done(function (result) {
                    document.getElementById("addProjectStatus").innerHTML = "";
                    $("#addProjectStatus").append("Project added");
                    localStorage.setItem("dailyLogProject" + document.getElementById("project").value, "" + result);
                    Excel.run(function (ctx) {
                        var tables = ctx.workbook.tables;

                        var project = document.getElementById("project").value;
                        var values = [[project]];
                        var row = tables.getItem("Projects").rows.add(null, values);
                        row.load('index');
                        return ctx.sync().then(function () {
                            console.log(row.index);
                        });
                    }).catch(function (error) {
                        console.log("Error: " + error);
                        if (error instanceof OfficeExtension.Error) {
                            console.log("Debug info: " + JSON.stringify(error.debugInfo));
                        }
                    });



                })
        });

    };


})();
//Notice function needs to be in global namespace
