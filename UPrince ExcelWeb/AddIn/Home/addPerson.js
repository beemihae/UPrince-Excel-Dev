(function () {

    Office.initialize = function (reason) {
        app.initialize();
        var host = 'https://uprinceuatapi.azurewebsites.net'
        var userEmail = localStorage.getItem('email');

        //If you need to initialize something you can do so here.
        $(document).on("click", "#addPerson", function () {
            var dataEmail =
           {
               "id": null,
               "firstName": document.getElementById("firstName").value,
               "lastName": document.getElementById("lastName").value,
               "company": document.getElementById("company").value,
               "email": document.getElementById("email").value,
               "function": document.getElementById("function").value,
               "phoneNo": document.getElementById("phone").value,
               "comment": document.getElementById("comment").value,
               "userEmail": userEmail,
               "isTeamMember": false
           };
            var urlProject = host + "/api/UserInvolvedPerson/PostUserInvolvedPerson"
            $.ajax({
                type: "POST",
                url: urlProject,
                dataType: "json",
                contentType: "application/json; charset=utf-8",
                data: JSON.stringify(dataEmail),
            })
                .done(function (str) {
                    if (str.isSuccess == true) {
                        document.getElementById("addPersonStatus").innerHTML = "";
                        $("#addPersonStatus").append("Person added");
                        Excel.run(function (ctx) {
                            var tables = ctx.workbook.tables;
                            
                            var responsible = document.getElementById("lastName").value
                            
                            if (/*document.getElementById("company").value != null ||*/ document.getElementById("company").value != "") {
                                responsible = responsible + " (" + document.getElementById("company").value + ")";
                            };
                            app.showNotification(document.getElementById("company").value + str.involvedPersonId);

                            console.log(responsible);
                            localStorage.setItem('dailyLogUsers' + responsible, str.involvedPersonId);
                            var values = [[responsible]];
                            var row = tables.getItem("Responsible").rows.add(null, values);
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

                    }
                    else {
                        document.getElementById("addPersonStatus").innerHTML = "";
                        $("#addPersonStatus").append("Wrong email");

                    };
                })
        });

    };


})();
//Notice function needs to be in global namespace
