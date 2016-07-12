﻿/// <reference path="../App.js" />
(function () {
    "use strict";
    var qualityCriteriaId;
    var projectId;
    var ProductDescriptionId;
    var projectPage = '<div class="main-wrapper"> <header class="col-lg-12 col-md-12 col-sm-12 col-xs-12 header-top"> <div class="col-lg-12 col-md-12 col-sm-12 col-xs-12 no-padding full-height"> <div class="header-sub header-glyph full-height"> <p title="UPrince.Projects"> <span class="glyphicon glyphicon-folder-open" aria-hidden="true"></span> </div> <div class="header-sub h1-div"> <h1 class="roboto-light">Projects</h1> </div> <div class="header-sub" style="position:absolute;right:15px"><p class="fake-link" id="logOut" style="font-size:12px;font-weight: 100; vertical-align: middle"> Log Out</p> </div></div> </header> <section class="col-lg-12 col-md-12 col-sm-12 col-xs-12 modal-div relationship container no-padding"><div><input id="projectSearch"></div> <div id="listProjects" class="nav nav-pills nav-stacked"></div> </section>  </div>'
    var host = 'https://uprincecoredevapi.azurewebsites.net';
    var previous = 0;
    // The initialize function must be run each time a new page is loaded
    Office.initialize = function (reason) {
        $(document).ready(function () {
            app.initialize();

            accessUser();

            //after log in go to project page, clicking sign in button
            $(document).on("click", "#btnSignIn", function () {
                //app.showNotification(JSON.stringify(bowser, null, '    '));
                var x = document.getElementById("email");
                var email = x.elements[0].value; 
                if (email.length != 0) {
                    sessionStorage.setItem("email", email);
                    //window.location.href = "project-page.html"
                    document.getElementById("login").innerHTML = "";
                    document.body.style.backgroundColor = "white";
                    //var projectPage = '<div class="main-wrapper"> <header class="col-lg-12 col-md-12 col-sm-12 col-xs-12 header-top"> <div class="col-lg-12 col-md-12 col-sm-12 col-xs-12 no-padding full-height"> <div class="header-sub header-glyph full-height"> <p  title="UPrince.Projects"> <span class="glyphicon glyphicon-folder-open" aria-hidden="true"></span></p> </div> <div class="header-sub h1-div"> <h1 class="roboto-light">Projects</h1> </div><div class="header-sub" style="position:absolute;right:15px"><p class="fake-link" id="logOut" style="font-size:12px;font-weight: 100; vertical-align: middle"> Log Out</p> </div> </div> </header> <section class="col-lg-12 col-md-12 col-sm-12 col-xs-12 modal-div relationship container no-padding"> <div id="listProjects" class="nav nav-pills nav-stacked"></div> </section> </div>'
                    $("#project-page").append(projectPage);
                    //window.location = "../project-page.html";
                    loadListProjects("");
                }
                else {
                    app.showNotification('Please enter login')
                }
            });

            //after log in go to project page, enter in emailfield
            $(document).submit("#email", function (event) {
                document.body.style.backgroundColor = "white";
                var x = document.getElementById("email");
                var email = x.elements[0].value;
                sessionStorage.setItem("email", email);
                //window.location.href = "project-page.html"
                document.getElementById("login").innerHTML = "";
                //var projectPage = '<div class="main-wrapper"> <header class="col-lg-12 col-md-12 col-sm-12 col-xs-12 header-top"> <div class="col-lg-12 col-md-12 col-sm-12 col-xs-12 no-padding full-height"> <div class="header-sub header-glyph full-height"> <p class="fake-link" title="UPrince.Projects"> <span class="glyphicon glyphicon-folder-open" aria-hidden="true"></span></p> </div> <div class="header-sub h1-div"> <h1 class="roboto-light">Projects</h1> </div> </div> </header> <section class="col-lg-12 col-md-12 col-sm-12 col-xs-12 modal-div relationship container no-padding"> <div id="listProjects" class="nav nav-pills nav-stacked"></div> </section> </div>'

                $("#project-page").append(projectPage);
                window.onload = loadListProjects("");
            })

            //go to product description page, after clicking a project
            $(document).on('click', "#listProjects li", function () {
                var projectId = $(this).attr("id");
                sessionStorage.setItem("projectId", projectId);
                var projectName = document.getElementById(projectId).innerHTML;
                //alert(projectName);
                sessionStorage.setItem('projectName', projectName);
                //window.location.href = "../project-menu.html"
                document.getElementById("project-page").innerHTML = "";
                var projectMenu = '<div class=main-wrapper><header class="col-sm-12 col-lg-12 col-md-12 col-xs-12 header-top"><div class="col-sm-12 col-lg-12 col-md-12 col-xs-12 no-padding full-height"><div class="header-sub full-height header-glyph"id=link-project-page><p title=UPrince.Projects><span class="glyphicon glyphicon-folder-open"aria-hidden=true></span></div><div class="header-sub h1-div"><h1 class=roboto-light id=projecName></h1></div></div></header><section class="col-sm-12 col-lg-12 col-md-12 col-xs-12 no-padding container modal-div relationship"><div class="col-sm-12 row-projects"id=Refresh><span class="glyphicon glyphicon-share-alt"></span>Refresh</div><div class="col-sm-12 row-projects"id=Publish><span class="glyphicon glyphicon-share-alt"></span>Publish</div><div class="col-sm-12 row-projects"id=createSheet><span class="glyphicon glyphicon-share-alt"></span>Create Sheet</div><div class="col-sm-12 row-projects"><input class=form-control id=filter> <span class="glyphicon form-control-filter glyphicon-filter"aria-hidden=true></span></div><div class="col-md-12 no-padding panel-group"><div class="panel panel-default"id=Cmdb><div class=panel-heading><h4 class=panel-title>CMDB</h4></div></div><div class="panel panel-default"id=DailyLog><div class=panel-heading><h4 class=panel-title><span class="fl-n icon-icon_DailyLog"></span> Daily Log</h4></div></div><div class="panel panel-default"id=IssueRegister><div class=panel-heading><h4 class=panel-title><span class="fl-n glyphicons cargo icon-icon_IssueRegister"></span> Issue Register</h4></div></div><div class="panel panel-default"id=LessonLog><div class=panel-heading><h4 class=panel-title><span class="fl-n glyphicons cargo icon-icon_LessonsLog"></span> Lesson Log</h4></div></div><div class="panel panel-default"id=Plans><div class=panel-heading><h4 class=panel-title><span class="fl-n glyphicons popup-menu-icon stats"></span> Plans</h4></div></div><div class="panel panel-default"id=ProductDescriptions><div class=panel-heading><h4 class=panel-title><span class="fl-n glyphicons popup-menu-icon stats"></span> Product Descriptions</h4></div></div><div class="panel panel-default"id=QualityRegister><div class=panel-heading><h4 class=panel-title><span class="fl-n icon-icon_QualityRegister"></span> Quality Register</h4></div></div><div class="panel panel-default"id=Reports><div class=panel-heading><h4 class=panel-title><span class="fl-n glyphicons popup-menu-icon stats"></span> Reports</h4></div></div><div class="panel panel-default"id=RiskRegister><div class=panel-heading><h4 class=panel-title><span class="fl-n icon-icon_RiskRegister"></span> Risk Register</h4></div></div></div></section></div>';
                $('#product-description-page').append(projectMenu);
                $('#projectName').append(projectName);
                //$(document).find('#saveBt').prop('disabled', false);
                //loadList()
                //$("#listProjects").append('<ul id="listProjects" class="nav nav-pills nav-stacked"><li id="productDescription"><a href="#">Product Description</a></li><li id="issueRegister"><a href="#">Issue Register</a></li><li id="dailyLog"><a href="#">Daily Log</a></li></ul>');

            });

            //click a product description, and opens a prod description
            /*$(document).on("click", 'ul li', function (e) {
                e.stopPropagation();
                sessionStorage.setItem('productDescriptionId', $(this).attr('id'));
                var div = $("<div>");
                Office.context.document.setSelectedDataAsync(div.html(), {
                    coercionType: "html"
                }, testForSuccess);
                getProductDescription();
            });*/


            //go back from prod descrp page to project page
            $(document).on("click", "#link-project-page", function () {
                document.getElementById("product-description-page").innerHTML = "";
                document.getElementById("login").innerHTML = "";
                $("#project-page").append(projectPage);
                loadListProjects("");
                document.getElementById("product-description-page").innerHTML = "";
            });

            //log out function, forget the email and go back to log out screen
            $(document).on('click', "#logOut", function () {
                document.getElementById("product-description-page").innerHTML = "";
                document.getElementById("login").innerHTML = "";
                document.getElementById("project-page").innerHTML = "";
                document.body.style.backgroundColor = "black";
                var loginScreen = ' <div class="top-content"> <div class="inner-bg"> <div class="container" style="background-color:black"> <div class="row"> <div class="col-sm-8 col-sm-offset-2 text"> <svg version="1.1" id="Layer_1" xmlns="http://www.w3.org/2000/svg" xmlns:xlink="http://www.w3.org/1999/xlink" x="0px" y="0px" width="80%" viewBox="0 0 1000 623" style="enable-background:new 0 0 1000 623;" xml:space="preserve"><path class="st1" d="M489.6,106.2c3.4-5.9,6.8-11.8,10.4-18c-0.5-0.4-1.1-0.7-1.5-1.1c-6.3-6.2-8.5-13.7-6.1-22.1c2.4-8.6,8.3-13.9,17-15.6c14.9-2.9,27.1,9.4,26.1,23.1c-0.4,5.3-2.4,9.8-6,13.7c-1.4,1.5-1.4,1.5-0.5,3.1c2.9,5,5.8,10,8.7,15.1c0.6,1.1,1.3,1.6,2.6,1.3c9.7-1.7,21.7,4.5,24.5,17.3c0.1,0.6,0.2,1.1,0.4,1.8c3.5,0,7,0,10.6,0c0.2-0.6,0.4-1.3,0.6-2c2.2-9.9,10.6-17.1,20.8-17.4c13.7-0.4,23.5,11.1,22.8,23.4c-0.6,10.7-9.9,21.3-23.1,20.7c-1.9-0.1-3.9-0.6-5.8-1.1c-0.9-0.2-1.5-0.2-2,0.7c-3.1,5.3-6.1,10.5-9.3,15.9c10,7.9,13.4,17.6,7.7,29.3c-4.1,8.5-13.2,12.5-21.2,11.7c-10.7-1.1-17-7.1-20.3-19.4c-3.4-0.1-6.9-0.1-10.5,0c-0.2,0.7-0.4,1.4-0.5,2c-2.2,9.1-8,14.7-17,17c-2,0.5-2,0.5-2.1,2.5c0,3.3,0,6.5,0,9.8c0,0.6,0.1,1.3,0.1,2.1c7.6,1.4,13.5,5.2,17.1,12.1c2.5,4.8,3,10,1.7,15.2c-2.3,9.7-11.5,17.6-23.5,16.6c-9.9-0.8-18.7-9.3-19.8-19.3c-1.2-11,5.4-22.3,19.3-24.8c0.1-0.6,0.2-1.2,0.2-1.8c0-3.4,0-6.8,0-10.2c0-1.9,0-1.9-1.8-2.3c-9.3-2.3-15.2-8-17.4-17.4c-0.1-0.5-0.3-1-0.5-1.6c-4.1,0-8.2,0-12.4,0c-1,6.4-4,11.5-9.2,15.3c-3.8,2.8-8.1,4.1-12.8,4.1c-9.3,0-17.7-6-20.8-14.8c-1.9-5.5-1.7-11,0.7-16.3c2.3-5.3,6.4-8.9,11.7-11.4c-3.7-5.3-7.3-10.4-10.9-15.6c-7.5,2.7-14.6,2.2-21.2-2.3c-5-3.5-8.1-8.3-9.1-14.3c-2.1-12.4,6-22.8,16.8-25.3c10.8-2.5,24.3,3.8,27,18.8c3.5,0,7,0,10.7,0c1.1-6.5,4.1-11.9,9.7-15.6C477,105.6,483,104.7,489.6,106.2z M539.8,149.6c-0.1,0.1-0.3,0.3-0.4,0.5c-3.4,5.5-6.7,11-10.1,16.5c-0.6,1-0.3,1.7,0.4,2.5c2.9,3.1,4.7,6.7,5.6,10.9c0.1,0.5,0.3,1.1,0.4,1.6c3.6,0,7,0,10.6,0c0.2-0.7,0.3-1.2,0.4-1.8c2.7-12.7,15.2-20,27-17c1.7,0.4,1.7,0.4,2.7-1.2c2.8-4.8,5.6-9.6,8.4-14.5c0.2-0.4,0.3-0.8,0.6-1.3c-2.6-1.9-4.9-4.1-6.3-6.9c-1.4-2.8-2.3-5.7-3.5-8.6c-3.4,0-6.8,0-10.2,0C563.2,142.8,552.6,150.9,539.8,149.6z M499.1,167.4c-4.6-5.8-9-11.4-13.3-17c-0.7-0.8-1.5-0.8-2.3-0.8c-10.2-0.3-18.6-7.1-21.1-17c-0.2-0.8-0.4-1.6-0.6-2.5c-3.6,0-7,0-10.5,0c-0.6,2-1,4-1.9,5.9c-0.8,1.8-1.9,3.6-3.1,5.2c-1.2,1.6-2.7,2.9-4.2,4.4c0.5,0.7,0.9,1.4,1.4,2.1c2.9,4.1,5.8,8.1,8.6,12.3c1,1.6,2.1,2.1,4.1,2.1c11.4-0.3,20.4,7,22.7,18.2c0.1,0.4,0.3,0.8,0.4,1.2c4.1,0,8.2,0,12.4,0C492.5,176,494.8,171.3,499.1,167.4z M521.3,130.1c-5.2,0-10.3,0-15.4,0c-0.2,0.6-0.4,1.1-0.5,1.5c-1.6,7.7-6,13-13.1,16.2c-0.5,0.2-0.9,0.5-1.6,0.9c4.2,5.3,8.3,10.5,12.4,15.7c7.2-3.5,14.3-3.4,21.4,0.3c3.5-5.7,6.8-11.2,10.2-16.9C527,144.2,522.7,138.3,521.3,130.1z M440,184.1c0,9.4,7.6,17.1,17,17.1c9.5,0,17.3-7.7,17.2-17.1c-0.1-10.9-9.3-17.2-17.1-17.2C448.3,166.9,439.8,174.3,440,184.1z M513.5,259.2c8.3,0.2,17.2-6.4,17.2-17.1c0-10.2-8.2-17.1-17.2-17.1c-9.4,0-17,7.6-17,17C496.4,251.6,504,259.2,513.5,259.2z M428.9,144.6c7.7,0.4,17.2-5.7,17.3-16.9c0.1-10.3-8.2-17.2-17-17.3c-9.5,0-17.2,7.6-17.2,17.1C412,137,419.5,144.6,428.9,144.6z M597.9,144.6c7.9,0.4,17.2-6,17.3-17.1c0.1-9.3-7.7-17-17.1-17.1c-9.5,0-17.2,7.6-17.2,17.1C580.9,136.9,588.5,144.6,597.9,144.6z M543.3,144.6c8.3,0.3,17.1-6.6,17.2-17.1c0-9.3-7.7-17.1-17-17.1c-9.5,0-17.1,7.6-17.3,17C526.1,136.8,534,144.8,543.3,144.6z M500.9,127.7c0.4-8.7-7-17.2-17-17.3c-9.4,0-17.2,7.7-17.2,17.1c0,9.4,7.6,17,17,17.1C493.1,144.8,501.1,136.8,500.9,127.7z M585.4,184c0.2-9.7-8.2-17.1-17.2-17.1c-7.8,0-16.9,6.2-17,17.1c-0.1,9.4,7.7,17.2,17.1,17.2C577.7,201.2,585.4,193.5,585.4,184z M513.6,88.1c8.5,0.1,17.1-6.6,17.1-17.2c-0.1-10-8-17-17.3-17.1c-9.3,0-17,7.8-17,17.1C496.5,80.5,504.1,88.1,513.6,88.1z M533.9,107.5c-3.3-5.8-6.5-11.3-9.8-17c-6.7,3.3-13.3,3.3-20.1,0.5c-3.3,5.8-6.5,11.3-9.8,17c3.3,2,5.8,4.2,7.8,7.2c2,2.9,3,6.2,3.8,9.7c5.2,0,10.3,0,15.5,0C522.6,117.1,526.7,111.3,533.9,107.5z" /><path class="st1" d="M184.8,309.7c0,19.3,0.6,38.6-0.1,57.9c-0.9,24.1-10.2,44.5-29.4,59.7c-10.1,8-21.8,12.9-34.4,15.4c-10,2-20.2,2.4-30.3,1.3c-19.2-2.1-36.2-9.1-50-23c-10.9-11-17.4-24.3-20.6-39.3c-1.5-7-1.9-14.1-1.9-21.2c0-35.9,0-71.7,0-107.6c0-1.2,0-2.3,0.2-3.5c0.4-3.5,2.9-6.2,6.3-6.8c1.1-0.2,2.3-0.3,3.5-0.3c10.2,0,20.5,0,30.7,0c0.4,0,0.7,0,1.1,0c5.4,0.2,8.4,3.3,8.4,8.7c0,22.5,0,45,0,67.5c0,14.8,0,29.6,0,44.4c0,4.7,0.6,9.4,2.1,13.9c3.7,11.4,11.6,18.1,23.3,20.2c6.1,1.1,12.2,1,18.2-0.5c12-3,19-10.8,21.9-22.6c1-4.2,1.3-8.6,1.3-13c0-36.6,0-73.2,0-109.7c0-1.2,0.1-2.5,0.4-3.6c1-2.9,3.1-4.7,6.2-5.2c0.9-0.2,1.9-0.2,2.8-0.2c10.4,0,20.8,0,31.1,0c1.1,0,2.2,0.1,3.2,0.3c3.4,0.7,5.9,3.6,6.1,7c0.1,0.8,0,1.6,0,2.4C184.9,271.2,184.9,290.4,184.8,309.7C184.9,309.7,184.8,309.7,184.8,309.7z M24.3,310.5c0,16.8,0,33.5,0,50.3c0,7.6,0.6,15.1,2.4,22.4c5,20.1,16.3,35.5,34.6,45.5c9.5,5.2,19.8,8,30.7,9c11.4,1,22.8,0.4,33.8-2.8c22.4-6.4,38.4-20.1,47.2-41.9c4.1-10,5.6-20.5,5.6-31.3c0-36.4,0-72.9,0-109.3c0-0.5,0-1,0-1.5c-0.1-1.9-0.6-2.4-2.5-2.5c-0.5,0-1,0-1.5,0c-9.8,0-19.6,0-29.4,0c-4,0-4,0-4,4c0,36.1,0,72.3,0,108.4c0,5.2-0.4,10.3-1.7,15.3c-3.4,12.8-11.1,21.7-23.9,25.7c-9.2,2.9-18.6,2.9-27.9,0.2c-10.4-3-17.7-9.6-22.1-19.5c-2.7-6.3-3.8-12.9-3.9-19.7c0-36.7,0-73.3,0-110c0-0.7,0-1.5,0-2.2c-0.1-1.7-0.5-2.1-2.2-2.3c-0.5,0-1,0-1.5,0c-9.9,0-19.7,0-29.6,0c-3.9,0-4.1,0.2-4.1,4C24.3,271.9,24.3,291.2,24.3,310.5z" /><path class="st1" d="M593.5,295.3c0.7-0.6,1.2-1,1.6-1.4c10-9.3,21.9-14.6,35.4-16.1c9-1,18-0.6,26.7,1.9c15.9,4.6,26.9,14.7,33.3,29.9c2.8,6.7,4.6,13.8,5.2,21c0.4,4.8,0.7,9.7,0.7,14.6c0.1,29.6,0,59.2,0,88.8c0,1.2,0,2.3-0.2,3.5c-0.6,3.4-3.2,5.8-6.7,6.1c-0.4,0-0.7,0.1-1.1,0.1c-9.9,0-19.7,0-29.6,0c-1,0-2.1-0.1-3-0.5c-3-1.2-4.7-3.5-4.9-6.8c0-0.7,0-1.5,0-2.2c0-29.8,0-59.5,0-89.3c0-3.9-0.3-7.7-1.3-11.4c-2.4-8.8-8.2-13.9-17.1-15.2c-11.9-1.9-22.5,1.2-32,8.6c-3.2,2.4-4.6,5.5-4.6,9.6c0.1,32.7-0.1,65.5,0.1,98.2c0,6.1-3.5,9.4-9.1,9.2c-9.1-0.3-18.3-0.1-27.4-0.1c-1,0-2.1,0-3-0.2c-3.1-0.7-5.2-3.4-5.6-6.8c-0.1-0.8,0-1.6,0-2.4c0-48.3,0-96.7,0-145c0-1.2,0-2.5,0.3-3.7c0.7-3.1,3-5.2,6.2-5.6c0.9-0.1,1.9-0.2,2.8-0.2c7.3,0,14.7,0,22,0c1.1,0,2.2,0.1,3.3,0.2c3,0.4,5,2.2,5.7,5.2c0.6,2.7,1.1,5.4,1.7,8.1C593.1,293.9,593.3,294.5,593.5,295.3z M557.2,361.8c0,24.1,0,48.2,0,72.3c0,3.3,0.1,3.3,3.2,3.3c8.7,0,17.4,0,26.1,0c0.6,0,1.2,0,1.7-0.1c0.8-0.1,1.3-0.5,1.4-1.4c0-0.9,0.1-1.7,0.1-2.6c0-32.5,0-65,0-97.5c0-5.6,2-10,6.3-13.5c10.9-8.8,23.3-12.5,37.2-10.5c11.8,1.7,19.8,9,22.7,20.6c1,4,1.4,8.2,1.4,12.3c0,29.8,0,59.7,0,89.5c0,3.1,0.1,3.2,3.2,3.2c8.6,0,17.3,0,25.9,0c3.7,0,3.7,0,3.7-3.7c0-27.6,0-55.2,0-82.7c0-5.3-0.1-10.6-0.3-15.9c-0.4-7.7-1.9-15.1-4.7-22.2c-4.1-10.3-10.8-18.5-20.8-23.6c-5.5-2.9-11.5-4.5-17.7-5.1c-9.1-1-18.1-0.4-26.9,2.4c-8.9,2.9-16.5,7.8-22.9,14.6c-1.6,1.7-3.5,2.3-5.7,1.7c-2.2-0.6-3.2-2.3-3.6-4.5c0-0.2-0.1-0.4-0.1-0.6c-0.7-3.3-1.4-6.7-2.1-10c-0.2-1.1-0.8-1.5-1.8-1.5c-0.5,0-1,0-1.5,0c-7.1,0-14.2,0-21.3,0c-3.5,0-3.5,0-3.5,3.5C557.2,313.7,557.2,337.8,557.2,361.8z" /><path class="st1" d="M874.3,376.2c1.4,11.8,6.3,20.9,16.2,27c6.2,3.8,13.1,5.3,20.3,5.6c10.3,0.4,19.5-2.4,27.4-9.1c3.2-2.7,6.9-2.7,10.2-0.2c6.8,5.2,13.6,10.4,20.3,15.7c4.1,3.2,4.3,8.4,0.4,12.1c-3.5,3.2-7.3,6-11.4,8.4c-9.7,5.8-20.1,9.2-31.3,10.7c-14.1,1.8-28.1,1.1-41.8-3.2c-17.5-5.4-31.9-15.2-42.4-30.3c-7.1-10.2-11.2-21.6-12.4-33.9c-0.3-3-0.4-5.9-0.5-8.9c0-7.1-0.2-14.2,0.1-21.3c0.8-17.9,6.4-34,18.4-47.6c9.9-11.2,22.3-18.4,36.8-22c17-4.2,34-3.9,50.5,2.1c23.6,8.6,38.2,25.5,44.7,49.6c1.9,6.9,2.7,13.9,2.7,21.1c0,5.3-0.2,10.6,0,15.9c0.2,4.9-3.5,8.6-8.5,8.6c-32.3-0.1-64.6,0-96.9,0C876.5,376.2,875.6,376.2,874.3,376.2z M924.2,369.9c16.1,0,32.2,0,48.3,0c0.8,0,1.6,0,2.4-0.1c0.9-0.1,1.4-0.6,1.5-1.5c0-0.5,0.1-1,0.1-1.5c0-4.9,0-9.7,0-14.6c0-7.1-0.8-14.1-2.8-20.9c-6.4-22.7-20.7-38.1-43.3-45.2c-16.6-5.2-33.4-4.8-49.9,0.7c-13.6,4.5-24.5,12.7-32.7,24.5c-7.5,11-11.3,23.3-11.9,36.6c-0.4,7.1-0.2,14.2-0.2,21.3c0,6.2,0.5,12.3,2.1,18.3c5.4,20.6,18,35.5,37,44.8c19.1,9.4,39.3,10.6,59.8,5.9c10.8-2.5,20.4-7.4,28.9-14.6c2.7-2.3,2.7-2.6,0-4.7c-6-4.7-12.1-9.3-18.1-14c-2-1.5-2.1-1.5-4.1,0.1c-4.1,3.2-8.5,5.9-13.4,7.5c-13,4.1-25.8,3.2-38.1-2.7c-5.4-2.6-9.9-6.4-13.5-11.1c-5-6.5-7.5-14-8.4-22.1c-0.4-3.3,2.1-6.3,5.5-6.7c0.9-0.1,1.9-0.1,2.8-0.1C892.3,369.9,908.3,369.9,924.2,369.9z" /><path class="st1" d="M705.4,361.7c0-4.3-0.2-8.6,0-12.8c0.8-15.9,6.4-30,16.3-42.3c11.8-14.7,27.1-24,45.4-28c20.8-4.6,40.9-2.3,60,7.6c1.7,0.9,3.3,1.8,4.9,2.8c3.8,2.5,4.9,6.5,2.6,10.4c-4,6.9-8.1,13.7-12.4,20.5c-2.3,3.7-6.3,4.5-10.3,2.3c-4.5-2.5-9.1-4.4-14.1-5.5c-7.6-1.6-15.1-1.9-22.6,0.4c-13.3,4.2-21.1,13.6-24.2,26.9c-0.7,2.9-0.9,6-0.9,9.1c-0.1,6.5-0.3,12.9,0,19.4c0.5,10.4,4.3,19.5,12.1,26.6c6.4,5.8,14,8.6,22.6,9c9.3,0.4,18-1.9,26.1-6.3c1.4-0.7,3-1.4,4.5-1.4c2.9-0.1,5.2,1.2,6.7,3.7c3.9,6.7,7.8,13.4,11.5,20.2c2.3,4.2,1,8.6-3.1,11.1c-9.6,5.7-19.9,9.4-31,10.9c-20,2.7-39-0.4-56.4-10.7c-18.8-11.2-31-27.4-36.1-48.8c-1.3-5.4-1.7-10.9-1.8-16.4c0-2.9,0-5.8,0-8.7C705.3,361.7,705.3,361.7,705.4,361.7z M711.4,361.8c0,0,0.1,0,0.1,0c0,2.8,0,5.5,0,8.3c0.1,5.4,0.6,10.9,1.9,16.2c5.4,21.5,18.4,36.9,38,46.7c14,7,28.9,9.1,44.4,7.5c10.8-1.1,21-4.4,30.4-9.8c2.6-1.5,2.7-1.9,1.2-4.4c-3.3-5.8-6.6-11.6-9.9-17.4c-1.5-2.6-1.5-2.6-4.1-1.2c-10.3,5.5-21.3,7.9-33,6.4c-9-1.2-16.9-4.8-23.3-11.1c-8.7-8.5-12.8-19-13.3-31c-0.2-6.2,0-12.5-0.1-18.7c0-3.9,0.4-7.7,1.3-11.4c4.1-15.8,13.7-26.5,29.6-31c11-3.1,21.9-1.8,32.4,2.2c2.6,1,5.1,2.4,7.6,3.6c1.5,0.7,1.7,0.7,2.6-0.7c0.2-0.2,0.3-0.5,0.5-0.7c3.5-5.8,7-11.7,10.6-17.5c0.2-0.3,0.4-0.6,0.6-0.9c0.7-1.3,0.6-1.7-0.6-2.5c-0.9-0.6-1.8-1.1-2.8-1.7c-18.4-10-37.8-12.2-58-7.5c-17.7,4.1-32.1,13.4-42.9,28.1c-7.8,10.6-12.2,22.5-12.9,35.7C711.5,353.1,711.5,357.4,711.4,361.8z" /><path class="st1" d="M256.9,376.4c-0.1,1.2-0.1,2.2-0.1,3.1c0,18.1,0,36.3,0,54.4c0,0.9,0,1.7,0,2.6c-0.2,4-3.2,6.9-7.2,7.1c-0.6,0-1.2,0-1.7,0c-10.7,0-21.3,0-32,0c-0.8,0-1.6,0-2.4-0.1c-3.6-0.4-6.3-3.3-6.6-7c-0.1-0.7,0-1.5,0-2.2c0-60.5,0-121.1,0-181.6c0-1.4,0.1-2.8,0.4-4.1c0.9-3.1,3.7-5.2,6.9-5.3c0.8,0,1.6,0,2.4,0c27.7,0,55.3,0.1,83,0c19.1-0.1,34.4,7.7,46.5,22.1c8.1,9.6,13,20.8,14.6,33.2c2.1,16.1,0,31.6-8,45.9c-9.7,17.2-24.4,27.5-43.8,30.9c-3.8,0.7-7.8,0.9-11.7,0.9c-12.4,0.1-24.8,0-37.2,0C259,376.4,258.1,376.4,256.9,376.4z M213.1,343.5c0,30,0,60,0,89.9c0,0.7,0,1.3,0,2c0.1,1.6,0.4,1.9,2,2c0.6,0,1.2,0,1.7,0c9.9,0,19.9,0,29.8,0c3.8,0,3.8,0,3.8-3.9c0-18.4,0-36.7,0-55.1c0-0.9,0-1.7,0.1-2.6c0.2-2.4,1.5-4.1,3.6-5c1.1-0.5,2.4-0.6,3.6-0.6c13.1,0,26.3,0,39.4-0.1c3.8,0,7.6-0.2,11.3-0.9c17.8-3.3,31-13.2,39.5-29.2c6.4-12.1,8-25.2,6.6-38.7c-1.1-11.2-5.2-21.4-12.2-30.3c-11-13.9-25.1-21.5-43.2-21.4c-27.4,0.2-54.7,0-82.1,0c-4.2,0-4,0-4,3.9C213.1,283.6,213.1,313.5,213.1,343.5z" /><path class="st1" d="M407.1,293c0.5-0.4,0.8-0.6,1.1-0.9c8.3-7.9,18.1-12.4,29.4-13.9c13.6-1.8,26.3,0.8,38,8c2.8,1.7,4.4,4.2,4.2,7.6c-0.1,1.3-0.6,2.7-1.3,3.9c-4.4,7-8.8,14-13.3,21c-2,3-5.2,4-8.8,2.9c-0.6-0.2-1.1-0.4-1.6-0.7c-9-4.8-18.4-4.6-27.9-1.6c-4.7,1.5-8.9,4.1-12.8,7.3c-2.8,2.3-4.1,5.2-4.1,8.8c0,2.1,0,4.2,0,6.3c0,31,0,62,0,93c0,0.9,0,1.8-0.1,2.6c-0.5,3.8-3.4,6.5-7.2,6.5c-4.1,0.1-8.3,0-12.4,0c-6,0-11.9,0-17.9,0c-3.7,0-6.5-2.2-7.3-5.7c-0.2-1-0.2-2-0.2-3c0-41.8,0-83.6,0-125.4c0-7.1,0-14.2,0-21.3c0-5.3,2.7-8.2,8.1-8.4c4.8-0.1,9.6,0,14.4,0c3.1,0,6.2-0.1,9.4,0c1.3,0,2.6,0.2,3.8,0.6c2.5,0.8,4,2.5,4.6,5.1C405.8,287.9,406.4,290.2,407.1,293z M371.2,361.7c0,24.3,0,48.6,0,72.9c0,2.6,0.1,2.7,2.7,2.7c0.3,0,0.6,0,0.9,0c8.4,0,16.8,0,25.3,0c3.6,0,3.6,0,3.6-3.6c0-32.4,0-64.7,0-97.1c0-0.8,0-1.6,0-2.4c0.3-4.5,1.8-8.5,5.2-11.6c12.3-11.6,33-16,48.3-7.8c2,1.1,2.2,0.9,3.5-1c3.8-6,7.6-12,11.4-18c1.9-3.1,1.9-3.4-1.2-5.2c-9.2-5.4-19.1-7.5-29.7-6.6c-11.7,1-21.7,5.3-29.9,13.8c-0.5,0.5-1,1.1-1.6,1.5c-3,2.2-7,0.9-8.1-2.6c-0.2-0.6-0.3-1.3-0.5-1.9c-0.6-2.3-1.1-4.6-1.8-7c-0.4-1.6-0.6-1.7-2.2-1.8c-0.4,0-0.9,0-1.3,0c-7,0-14.1,0-21.1,0c-3.5,0-3.5,0-3.5,3.5C371.2,313.7,371.2,337.7,371.2,361.7z" /><path class="st1" d="M497.3,443.4c-4,0.5-7.5-2.5-7.8-6.4c-0.1-0.7,0-1.3,0-2c0-48.8,0-97.7,0-146.5c0-1.3,0.1-2.6,0.5-3.8c1-3.1,4-4.9,7.3-4.5c0,1.6,0,3.3,0,4.9c0,0.3-0.1,0.7-0.3,0.8c-1.6,0.9-1.2,2.3-1.2,3.7c0,48.2,0,96.3,0,144.5c0,1.3-0.4,2.8,1.2,3.7c0.2,0.1,0.3,0.5,0.3,0.8C497.3,440.1,497.3,441.8,497.3,443.4z" /><path class="st1" d="M526.7,443.5c0-1.7,0-3.4,0-5c0-0.3,0.2-0.8,0.4-0.9c1.2-0.7,1.1-1.8,1.1-2.9c0-48.6,0-97.3,0-145.9c0-1,0.1-2-1-2.7c-0.3-0.2-0.5-0.7-0.6-1.1c-0.1-1.6,0-3.2,0-5.1c1.1,0.1,1.9,0.1,2.7,0.3c2.9,0.9,4.9,3.5,5.1,6.6c0.1,0.7,0,1.5,0,2.2c0,48.5,0,97,0,145.5c0,1.4-0.1,2.8-0.5,4.1C533.1,441.7,530.5,443.6,526.7,443.5z" /><path class="st1" d="M398.9,555.8c0-0.9,0-1.6,0-2.3c0-13.7,0-27.4,0-41.1c0-0.4,0-0.9,0-1.3c0-1,0.6-1.5,1.5-1.4c5.2,0.1,10.4-0.1,15.6,0.4c4.5,0.4,7.8,2.7,8.9,7.4c1.1,4.8,0,8.9-4.1,11.9c-0.2,0.2-0.4,0.3-0.9,0.7c0.6,0.2,1,0.4,1.4,0.5c6.3,1.9,9.7,6.9,9.1,13.5c-0.5,6.3-5.3,11.5-11.6,11.9c-6.1,0.4-12.3,0.2-18.5,0.2C400,556.2,399.6,556,398.9,555.8z M402.1,553.4c0.5,0,0.9,0.1,1.4,0.1c4.2,0,8.4,0,12.6,0c0.9,0,1.9-0.2,2.8-0.4c5.3-1.1,8.3-4.8,8.3-10.3c0-5.5-2.8-9.2-8.2-10c-5.6-0.9-11.2-0.2-16.8-0.3C401.7,534.4,401.6,551.1,402.1,553.4z M402.2,529.8c4.5-0.1,9.1,0.7,13.7-0.6c4.5-1.3,6.6-4.6,6.2-9.4c-0.3-3.9-3-6.6-7.4-7.2c-0.6-0.1-1.3-0.1-1.9-0.1c-3,0-5.9,0-8.9,0c-0.6,0-1.1,0.1-1.6,0.2C401.6,514.5,401.7,527.7,402.2,529.8z" /><path class="st1" d="M765.8,516.9c0,0.8,0,1.6,0,2.4c0,11.1,0,22.2,0,33.3c0,0.4,0,0.9,0,1.3c0,2.2-0.3,2.4-2.6,1.9c-0.6-2.1-0.6-44.2,0.1-46c1.1-0.2,1.5,0.7,2,1.3c4.6,5.4,9.1,10.7,13.7,16.1c5.7,6.7,11.4,13.4,17,20.1c0.5,0.5,1,1,1.8,1.9c0.1-1.2,0.1-2,0.1-2.7c0-11.2,0-22.4,0-33.5c0-0.7,0-1.3,0-2c0-1,0.5-1.5,1.5-1.5c1,0,1.4,0.5,1.5,1.5c0,0.5,0,1,0,1.5c0,13.5,0,27,0,40.5c0,0.9-0.1,1.8-0.2,2.7c-1.3,0.4-1.8-0.5-2.3-1.2c-3.3-3.9-6.7-7.9-10-11.8c-7-8.3-14-16.5-21-24.8c-0.4-0.5-0.8-0.9-1.3-1.4C766.1,516.8,765.9,516.9,765.8,516.9z" /><path class="st1" d="M539.8,555.8c-1.6,0.4-2-0.6-2.5-1.2c-2.9-3.4-5.7-6.7-8.6-10.1c-7.4-8.7-14.8-17.5-22.2-26.2c-0.5-0.5-1-1.1-1.8-2c-0.2,3-0.1,5.5-0.1,8.1c0,2.6,0,5.2,0,7.8c0,2.6,0,5.2,0,7.8s0,5.2,0,7.8c0,2.6,0,5.2,0,8c-1,0-1.8,0-2.9,0c0-15.4,0-30.7,0-45.8c1.5-0.5,1.9,0.4,2.4,1c2.6,3,5.1,6,7.6,9c7.9,9.3,15.8,18.6,23.7,27.9c0.4,0.4,0.8,0.9,1.5,1.7c0-1.2,0-2,0-2.7c0-11.4,0-22.8,0-34.2c0-0.6,0-1.2,0-1.7c0-0.9,0.5-1.3,1.3-1.3c0.9,0,1.5,0.3,1.5,1.4c0,0.5,0,1,0,1.5c0,13.6,0,27.3,0,40.9C539.8,554.2,539.8,555,539.8,555.8z" /><path class="st1" d="M295.3,509.9c1.6,0,2,1,2.7,1.8c8,9.3,15.9,18.7,23.8,28c2.2,2.6,4.4,5.2,6.6,7.8c0.5,0.5,1,1,1.8,2c0.1-1.2,0.1-2,0.1-2.7c0-11.3,0-22.6,0-34c0-0.6,0-1.2,0-1.7c0-0.9,0.4-1.4,1.3-1.4c0.9,0,1.6,0.3,1.6,1.3c0,0.5,0,1,0,1.5c0,13.6,0,27.3,0,40.9c0,0.9,0.5,2.1-0.7,2.5c-1.1,0.3-1.5-0.9-2.1-1.6c-10.1-12-20.3-23.9-30.4-35.9c-0.5-0.6-1-1.2-2-1.7c0,1.8,0,3.6-0.1,5.4c0,1.8,0,3.6,0,5.4c0,1.9,0,3.8,0,5.7c0,1.9,0,3.8,0,5.7c0,1.9,0,3.8,0,5.7c0,1.8,0,3.6,0,5.4c0,1.9,0,3.8,0,5.8c-1,0-1.8,0-2.7,0C295.3,540.5,295.3,525.3,295.3,509.9z" /><path class="st1" d="M849.3,556.6c-1.1-0.9-1.7-1.5-2.3-2c-2.3,0.7-4.6,1.6-7,2c-12.7,2-24.8-5.6-25.2-20.5c-0.1-3.4-0.1-6.9,0.5-10.2c1.7-9.7,9.6-16.3,19.5-16.7c3.7-0.2,7.2,0.3,10.6,1.8c0.7,0.3,1.3,0.8,1.9,1.1c-1,2.2-1.1,2.3-2.9,1.5c-4.5-2.1-9.1-2.3-13.8-1c-6.4,1.8-10.4,6.2-12,12.6c-1.3,5.1-1.3,10.2-0.1,15.3c1.9,8.2,8.3,13.3,16.7,13.6c3.3,0.1,6.5-0.2,9.6-1.4c1.3-0.5,1.9-1.3,1.9-2.8c-0.1-3.8,0-7.7,0-11.5c0-0.6-0.1-1.1-0.1-1.9c-0.9-0.1-1.8-0.2-2.6-0.2c-2.2,0-4.5,0-6.7-0.1c-0.6,0-1.4-0.3-1.6-0.7c-0.5-0.9,0.1-1.9,1.3-1.8c1.9,0,3.8,0,5.7,0c1.7,0,3.5,0,5.2,0c1.7,0,1.9,0.2,1.9,1.9c0.1,1.9,0,3.8,0,5.7c0,4.1,0,8.3,0,12.4C849.7,554.3,849.5,555.2,849.3,556.6z" /><path class="st1" d="M585.8,554.2c-3.6,1.9-7.3,2.7-11.2,2.6c-11.2-0.2-19.4-7.5-20.8-18.6c-0.5-4.5-0.6-9,0.5-13.4c2.2-9.1,9.7-15.3,19-15.7c3.7-0.2,7.2,0.2,10.6,1.7c0.5,0.2,0.9,0.4,1.4,0.6c0.7,0.4,1,0.9,0.5,1.6c-0.4,0.6-0.8,1.2-1.7,0.9c-0.4-0.2-0.8-0.4-1.2-0.5c-4.5-2-9.2-2.2-13.8-0.7c-5.9,1.8-9.8,5.8-11.5,11.8c-1.5,5.5-1.5,11.1-0.1,16.6c2,7.8,8.4,12.7,16.5,13c3.3,0.1,6.5-0.2,9.6-1.4c1.3-0.5,2.1-1.3,2-2.9c-0.1-3.8,0-7.5,0-11.3c0-0.6-0.1-1.3-0.2-2.2c-1.6,0-3.2,0-4.7,0c-1.6,0-3.2,0-4.8,0c-0.8,0-1.6-0.1-1.6-1.3c0-1.1,0.7-1.3,1.6-1.3c3.6,0,7.3,0,10.9,0c1.5,0,1.7,0.3,1.8,1.9c0,6.3,0,12.6,0,18.9c0,0.6,0.2,1.5-0.7,1.6c-0.4,0-0.9-0.5-1.3-0.9C586.3,554.9,586.2,554.6,585.8,554.2z" /><path class="st1" d="M666.6,556c-1.1,0-1.9,0-2.8,0c0-0.9-0.1-1.7-0.1-2.5c0-13.7,0-27.4,0-41.1c0-0.1,0-0.3,0-0.4c0-1.9,0.3-2.3,1.5-2.3c1.1,0,1.4,0.4,1.4,2.2c0,4.9,0,9.9,0,14.8c0,1.2,0,2.5,0,3.7c0,0.9,0.4,1.3,1.3,1.4c0.4,0,0.7,0,1.1,0c7.6,0,15.2,0,22.8,0c0.7,0,1.4-0.1,2.4-0.2c0-0.9,0-1.7,0-2.5c0-5.5,0-11,0-16.5c0-0.6,0-1.2,0.1-1.7c0.1-0.9,0.7-1.1,1.6-1.1c0.7,0,1.3,0.3,1.3,1.1c0.1,0.7,0.1,1.3,0.1,2c0,13.4,0,26.8,0,40.3c0,0.9-0.1,1.9-0.2,2.7c-2.5,0.8-2.9,0.5-2.9-1.8c0-5.7,0-11.3,0-17c0-0.8,0-1.6,0-2.3c-1.9-0.5-25.2-0.6-27.6-0.1c-0.1,3.4,0,7,0,10.5C666.6,548.6,666.6,552.2,666.6,556z" /><path class="st1" d="M750.3,555.8c-2.6,0.6-2.8,0.5-3.6-1.3c-1.3-2.9-2.5-5.7-3.7-8.6c-0.5-1.3-1.3-1.7-2.6-1.7c-7,0-13.9,0.1-20.9,0c-1.4,0-2.1,0.5-2.6,1.8c-1.2,2.9-2.4,5.7-3.7,8.6c-0.8,1.8-1.1,1.9-3.5,1.2c0.3-0.8,0.5-1.6,0.8-2.4c5.2-12,10.4-24,15.5-35.9c0.8-1.9,1.7-3.9,2.5-5.8c0.3-0.6,0.7-1.2,1-1.8c1.4,0.1,1.5,1.1,1.8,1.9c2.5,5.6,4.9,11.3,7.4,16.9c3.6,8.2,7.1,16.4,10.6,24.5C749.8,554,750,554.8,750.3,555.8z M730.3,516c-0.2,0-0.4,0-0.6,0c-3.5,8.3-7.1,16.6-10.6,25.3c7.5,0,14.6,0,22,0C737.4,532.7,733.9,524.3,730.3,516z" /><path class="st1" d="M386.4,555.6c-2.2,0.9-2.8,0.8-3.5-0.9c-1.3-2.8-2.5-5.7-3.7-8.6c-0.6-1.4-1.3-2-2.9-2c-6.9,0.1-13.8,0.1-20.7,0c-1.5,0-2.2,0.6-2.7,1.9c-1.2,2.9-2.4,5.7-3.7,8.6c-0.8,1.7-1.2,1.9-3.2,1.2c0.2-1.9,18.6-44.5,19.9-45.9c1,0.2,1.1,1.2,1.5,1.9c2.3,5.3,4.6,10.6,6.9,15.9c3.7,8.5,7.4,17,11.1,25.5C385.8,554,386.1,554.8,386.4,555.6z M366.3,515.8c-0.2,0-0.3,0-0.5,0c-3.6,8.4-7.1,16.8-10.8,25.5c7.5,0,14.6,0,22.1,0C373.5,532.6,369.9,524.2,366.3,515.8z" /><path class="st1" d="M867.4,512.4c0,6.5,0,12.7,0,19.1c0.8,0.1,1.5,0.2,2.3,0.2c4.4,0,8.9,0,13.3,0c0.6,0,1.2,0,1.7,0.1c0.7,0.1,0.9,0.6,0.9,1.2c0,0.7-0.3,1.1-1,1.2c-0.6,0-1.2,0-1.7,0c-4.3,0-8.6,0-12.8,0c-0.8,0-1.6,0-2.3,0c-0.6,1.7-0.7,15.1-0.3,19.2c0.7,0,1.5,0.1,2.2,0.1c6.9,0,13.8,0,20.7,0c0.6,0,1.2,0,1.7,0c0.8,0.1,1.1,0.6,1.1,1.3c0,0.7-0.4,1.1-1.1,1.1c-0.6,0.1-1.3,0.1-2,0.1c-7.6,0-15.2,0-22.9,0c-0.9,0-1.8-0.1-2.9-0.2c0-1-0.1-1.7-0.1-2.5c0-13.7,0-27.4,0-41.1c0-0.4,0-0.9,0-1.3c0-0.9,0.5-1.3,1.3-1.4c0.5,0,1,0,1.5,0c7.8,0,15.5,0,23.3,0c0.5,0,1,0,1.5,0c0.8,0.1,1.2,0.5,1.2,1.2c0,0.9-0.4,1.4-1.3,1.4c-0.6,0.1-1.2,0-1.7,0c-6.7,0-13.5,0-20.2,0C869.2,512.4,868.5,512.4,867.4,512.4z" /><path class="st1" d="M253.8,556c0-1,0-1.7,0-2.5c0-13.7,0-27.4,0-41.1c0-0.4,0-0.9,0-1.3c0-0.9,0.5-1.3,1.3-1.4c0.4,0,0.9,0,1.3,0c7.8,0,15.7,0,23.5,0c0.4,0,0.9,0,1.3,0c0.8,0,1.4,0.4,1.4,1.3c0,1-0.6,1.4-1.5,1.4c-0.6,0-1.2,0-1.7,0c-6.7,0-13.3,0-20,0c-0.8,0-1.6,0-2.6,0c-0.1,6.4,0,12.6-0.1,19.1c0.9,0.1,1.6,0.2,2.4,0.2c4.4,0,8.8,0,13.3,0c0.6,0,1.2,0,1.7,0.1c0.7,0.1,0.9,0.6,0.9,1.2c0,0.7-0.3,1.1-1,1.2c-0.6,0-1.2,0-1.7,0c-4.4,0-8.7,0-13.1,0c-0.7,0-1.4,0-2.1,0c-0.5,1.7-0.7,15.4-0.3,19.2c0.7,0,1.5,0.1,2.2,0.1c6.9,0,13.8,0,20.7,0c0.6,0,1.2,0,1.7,0c0.8,0.1,1.1,0.6,1.1,1.3c0,0.7-0.4,1.1-1.1,1.1c-0.6,0.1-1.3,0.1-2,0.1c-7.6,0-15.2,0-22.8,0C255.9,556.2,254.9,556.1,253.8,556z" /><path class="st1" d="M650.4,554.4c-2.5,1.6-5.3,2.1-8.1,2.4c-3.1,0.3-6.1-0.1-9-1c-7.7-2.5-12.3-7.8-13.9-15.6c-1-5.2-1-10.5,0.3-15.8c2.2-8.9,9.7-14.9,18.9-15.3c3.4-0.2,6.7,0.1,9.8,1.4c0.7,0.3,1.3,0.6,1.9,0.9c-0.5,2.3-0.8,2.5-2.7,1.8c-4.3-1.7-8.7-1.8-13-0.5c-7.3,2.2-11.2,7.5-12.3,14.9c-0.6,4.1-0.6,8.3,0.2,12.4c1.7,8.4,8.3,13.9,16.9,14.2c2.9,0.1,5.6-0.2,8.3-1.3C649.4,552,649.5,552.1,650.4,554.4z" /><path class="st1" d="M446.1,553.6c1.1,0,1.9,0,2.6,0c6.7,0,13.3,0,20,0c0.5,0,1,0,1.5,0c0.9,0,1.6,0.2,1.6,1.3c0,1.1-0.8,1.2-1.6,1.3c-0.3,0-0.6,0-0.9,0c-8,0-15.9,0-23.9,0c-2.3,0-2.5-0.2-2.5-2.4c0-13.9,0-27.8,0-41.8c0-1.9,0.3-2.3,1.5-2.4c0.8,0,1.4,0.3,1.5,1.2c0.1,0.8,0.1,1.6,0.1,2.4c0,12.4,0,24.8,0,37.2C445.9,551.3,446,552.3,446.1,553.6z" /><path class="st1" d="M485.8,555.9c-2.7,0.4-2.9,0.3-2.9-2c0-13.9,0-27.9,0-41.8c0-0.4,0-0.9,0.1-1.3c0.1-0.9,0.7-1.2,1.5-1.2c0.7,0,1.3,0.3,1.3,1.1c0.1,0.6,0.1,1.3,0.1,2c0,13.4,0,26.8,0,40.3C486,553.9,485.9,554.8,485.8,555.9z" /><path class="st1" d="M906.2,351.1c-9.9,0-19.7,0-29.6,0c-1,0-2,0-3-0.2c-3.3-0.5-5.4-2.9-5.1-6.2c0.3-3,1-6,1.7-9c0.9-3.8,2.4-7.3,4.4-10.6c4.1-6.8,9.9-11.7,17.5-14.3c8.5-2.9,17.2-3.3,25.9-1.1c11.4,2.9,18.8,10.4,23,21.3c1.5,3.9,2.4,7.9,3,12c0.1,0.8,0.1,1.6,0.1,2.4c-0.1,2.6-1.8,4.7-4.3,5.3c-1.1,0.2-2.3,0.3-3.5,0.3C926.2,351.1,916.2,351.1,906.2,351.1z M874.7,344.4c2.1,0.6,61,0.6,62.8,0.1c0.1-0.1,0.1-0.3,0.2-0.4c0-0.3,0-0.6,0-0.9c-0.8-5.3-2.2-10.3-4.9-15c-3.7-6.3-9.1-10.5-16.2-12.3c-5.2-1.4-10.5-1.5-15.8-0.7c-6,0.9-11.5,3-16,7.2c-5,4.6-7.6,10.5-9.1,17C875.2,341.1,875,342.8,874.7,344.4z" /><path class="st1" d="M250.5,311.6c0-6.9,0-13.8,0-20.7c0-0.9,0-1.7,0.1-2.6c0.5-3.1,2.9-5.2,6-5.3c0.4,0,0.9,0,1.3,0c8.9,0,17.9,0,26.8,0c3.1,0,6.2,0.4,9.2,1.3c8.2,2.6,13.7,8,16.3,16.2c2.3,7.5,2.3,15,0,22.5c-3,9.9-11.3,16.5-21.7,17.4c-1.9,0.2-3.8,0.3-5.6,0.3c-8.1,0-16.1,0-24.2,0c-0.9,0-1.7,0-2.6-0.1c-3.1-0.5-5.2-2.7-5.6-5.8c-0.1-0.7-0.1-1.4-0.1-2.2C250.5,325.6,250.5,318.6,250.5,311.6z M256.9,289.5c0,0.9-0.1,1.7-0.1,2.5c0,13.2,0,26.4,0,39.6c0,0.5,0,1,0,1.5c0,0.8,0.5,1.2,1.3,1.2c0.6,0,1.2,0,1.7,0c7.6,0,15.2,0,22.9,0c1.9,0,3.8-0.1,5.6-0.3c8-0.8,13.9-5.8,16.1-13.5c1.3-4.4,1.3-8.9,0.9-13.4c-0.9-10.2-8.1-17.4-18.3-17.8c-9.6-0.3-19.2-0.1-28.7-0.1C257.9,289.3,257.5,289.4,256.9,289.5z" /><circle class="st0" cx="513.3" cy="184" r="22.2" /></svg> </div> </div> <div class="row"> <div class="col-md-6 col-md-offset-3 form-box"> <div class="form-top"> <h3>Welcome to UPrince Word</h3> <p>The easiest way to edit PRINCE2 Project Descriptions in Word</p> <p>Please register at <a class="aWelcome" href="https://start.uprince.com">UPrince</a> before you can access your projects.</p> </div> <div class="form-bottom col-sm-12 "> <form id="email" onsubmit="submitLogin"> <div class="col-sm-9"> <input type="text" autofocus required placeholder="Email" name="email" class="form-username form-control" id="email"> </div> <div class="col-sm-3"> <!--<button id="btnSignIn" type="submit" class="btn" disabled>Sign in!</button>--> <button id="btnSignIn" type="button" class="btn">Sign in!</button> </div> </form> </div> </div> </div> </div> </div> </div>'
                $("#login").append(loginScreen);
                sessionStorage.setItem("email", "");

            });
            $(document).on('input', '#projectSearch', function () {
                var x = document.getElementById("projectSearch").value;
                var projectSearch;
                var update;
                if (x.length > 2) {
                    projectSearch = x;
                } else { projectSearch = ""; };
                //document.getElementById("name").innerHTML = projectSearch
                if ((previous == 1) && (x.length == 2)) { };
                document.getElementById("listProjects").innerHTML = "";
                loadListProjects(projectSearch);
                previous = x.length;
            });

            $(document).on("click", "#createSheet", function () {
                Excel.run(function (ctx) {
                    ctx.workbook.worksheets.add("Values");
                    return ctx.sync().then(function () {
                        app.showNotification("Success! Worksheet created with a random name.");
                    });
                }).catch(function (error) {
                    app.showNotification(error);
                });
            });

        });
    };

    function accessUser() {
        var code = getToken();
        //$("#code").append(code);
        //$("#status").append('');
        var url = "https://uprince-dev.pronovix.net/api/system/connect"
        var authorization = "Bearer " + code;

        //JQuery
        $.ajax({
            type: "POST",
            url: url,
            dataType: "json",
            //contentType: "application/json; charset=utf-8",
            headers: { "Authorization": authorization }
        })
          .done(function (str) {
              document.body.style.backgroundColor = "white";
              var email = str.user.mail;
              sessionStorage.setItem("email", email);
              //window.location.href = "project-page.html"
              document.getElementById("login").innerHTML = "";
              $("#project-page").append(projectPage);
              loadListProjects("");
          })
         .fail(function (jqXHR, textStatus, errorType) {
             $("#status").append(textStatus + ' ' + errorType);
         });

        //XHR
        /*var http = new XMLHttpRequest();
        http.open("POST", url, true);
        //Send the proper header information along with the request
        http.setRequestHeader("Content-type", "application/json; charset=utf-8");
        http.setRequestHeader("Authorization", authorization);
        http.setRequestHeader("dataType", "json");
        //http.setRequestHeader("Connection", "close");
        http.onreadystatechange = function () {//Call a function when the state changes.
            if (http.readyState == 4 && http.status == 200) {
                alert(http.responseText);
            };
        };
        http.send();
       */


    };


    function getToken() {
        var url = window.location.href;
        var startParam = url.indexOf('access_token');
        var start = url.indexOf('=', startParam) + 1;
        var eind = url.indexOf('&', start)
        return url.substring(start, eind);
    }

    //load projects in to projectpage from server
    function loadListProjects(projectName) {
        var email = sessionStorage.getItem('email')
        var dataEmail = {
            "customer": "",
            "email": email,
            "isFocused": {
                "customer": false,
                "title": false
            },
            "isRecycled": false,
            "orderField": "id",
            "sortOrder": "ASC",
            "status": {
                "Active": false,
                "All": true,
                "Closed": false,
                "New": false
            },
            "title": projectName,
            "toleranceStatus": {
                "All": true,
                "OutofTolerance": false,
                "Tolerancelimit": false,
                "WithinTolerance": false
            }
        };
        $.ajax({
            type: "POST",
            url: host + "/api/project/GetProjectList",
            dataType: "json",
            contentType: "application/json; charset=utf-8",
            data: JSON.stringify(dataEmail),
        })
          .done(function (str) {
              document.getElementById("listProjects").innerHTML = "";
              var test = str;
              var length = Object.keys(str).length;
              $("#listProjects").append('<ul id="listProjects" class="nav nav-pills nav-stacked">');
              for (var i = 0; i < length; i++) {
                  var dummy = "<li id='".concat(str[i].id, "'><a href='#'>", str[i].title, "</a></li>");
                  $("#listProjects").append(dummy);
              }
          })


    };

    //load all the product descrip from server
    function loadList() {
        var projectName = sessionStorage.getItem('projectName');
        $("#projectName").html(projectName);
        $("#listProductDescription").html('');
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
              organizeList(str);
          })
    };

    //lay out from prod descpr page
    function organizeList(str) {
        var length = str.length;
        var count = 0;
        var id = [];
        for (var i = 0; i < length; i++) {
            if (str[i].ParentId == null) {
                var dummy = '<li id="' + str[i].Id + '"><a href="javascript: void(0);" class="p-l-30">' + str[i].Title + '</a> ';
                if (str[i].HasChildren) {
                    dummy = dummy + '<ul id="L' + str[i].Id + '" class="nav nav-pills nav-stacked"></li>'
                } else {
                    dummy = dummy + '</li>'

                }
                $("#expList").append(dummy);
                count++;
                id.push("" + str[i].Id)
            }
        };
        while (count < length) {
            for (var i = 0; i < length; i++) {
                if ((id.indexOf("" + str[i].Id) == -1) && (id.indexOf("" + str[i].ParentId) != -1)) {
                    var dummy = '<li id="' + str[i].Id + '"><a href="javascript: void(0);" class="p-l-30">' + str[i].Title + '</a>';
                    //var dummy = '<li id="' + str[i].Id + '"><span  class="p-l-30 fake-link">' + str[i].Title + '</span>';

                    if (str[i].HasChildren) {
                        dummy = dummy + '<ul id="L' + str[i].Id + '" class="nav nav-pills nav-stacked"></li>'
                    } else {
                        dummy = dummy + '</li>'

                    }
                    $("#L" + str[i].ParentId).append(dummy);
                    count++;
                    id.push("" + str[i].Id)
                };
            };
        };
        //prepareList();

    };

    //expand and colllapse list, not in use
    function prepareList() {
        $('#expList').find('li:has(ul)')
          .click(function (event) {
              if (this == event.target) {
                  $(this).toggleClass('expanded');
                  $(this).children('ul').toggle('medium');
                  e.stopPropagation();
                  sessionStorage.setItem('productDescriptionId', $(this).attr('id'));
                  var div = $("<div>");
                  Office.context.document.setSelectedDataAsync(div.html(), {
                      coercionType: "html"
                  }, testForSuccess);
                  getProductDescription();
              }
              return false;
          })
          .addClass('collapsed')
          .children('ul').hide();

    };

    //test for completion of request
    function testForSuccess(asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            //app.showNotification('Error', asyncResult.error.message);
        }
    };

})();