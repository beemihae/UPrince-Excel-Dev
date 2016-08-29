(function () {

    Office.initialize = function (reason) {
        app.initialize();
        var host = 'https://uprinceuatapi.azurewebsites.net'
        var userEmail = localStorage.getItem('email');

        //If you need to initialize something you can do so here.
        $(document).on("click", "#addContext", function () {
            app.showNotification(document.getElementById("context").value)
            var dataEmail = {
                "id": null,
                "description": document.getElementById("context").value,
                "coreUserEmail": userEmail,

            };
            var urlProject = host + "/api/DailyLog/PostDailyLogContext";
            $.ajax({
                type: "POST",
                url: urlProject,
                dataType: "json",
                contentType: "application/json; charset=utf-8",
                data: JSON.stringify(dataEmail),
            })
                .done(function (result) {
                    document.getElementById("addContextStatus").innerHTML = "";
                    $("#addContextStatus").append("Context added");
                    localStorage.setItem("dailyLogContext" + document.getElementById("context").value, "" + result);
                    Excel.run(function (ctx) {
                        var tables = ctx.workbook.tables;

                        var project = document.getElementById("context").value;
                        var values = [[project]];
                        var row = tables.getItem("atContext").rows.add(null, values);
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
