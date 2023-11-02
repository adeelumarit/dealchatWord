// <reference path="messageread.js" />
//var baseURL = "https://localhost:44371"
var baseURL = "https://accountingsystemapi20231101183220.azurewebsites.net"

var app = angular.module('DealChat', ['ngMaterial', "ngRoute"], function () {


});




app.config(function ($routeProvider) {
    $routeProvider
        .when("/", {
            templateUrl: "/Templates/mainPage.html"
        })
        .when("/Signup", {
            templateUrl: "/Templates/Signup.html"
        })
        .when("/Login", {
            templateUrl: "/Templates/Login.html"



        })
        .when("/blue", {
            templateUrl: "blue.htm"
        });
});

app.controller('Signupctrl', function ($scope, $mdDialog, $mdToast, $log, $location,) {


    $scope.create_User = function () {
        ProgressLinearActive();
        var userObject = {
            name: $scope.user_name,
            email: $scope.user_email,
            password: $scope.user_password
            // Other properties of the user object
        };

        var settings = {
            "url": baseURL +"/api/Home/createUser",
            "method": "POST",
            "timeout": 0,
            "headers": {
                "Content-Type": "application/json"
            },
            "data": JSON.stringify({
                name: $scope.user_name,
                email: $scope.user_email,
                password: $scope.user_password
            }),
        };

        $.ajax(settings).done(function (response) {
            console.log(response);
            $location.path("/Login")
            ProgressLinearInActive();

        }).fail(function (error) {

            ProgressLinearInActive()
            console.log(error)
        });


    }
    function ProgressLinearActive() {
        $("#StartProgressLinear").show(function () {

            $("#ProgressBgDiv").show();
            $scope.ddeterminateValue = 15;
            $scope.showProgressLinear = false;
            if (!$scope.$$phase) {
                $scope.$apply();
            }
        });
    };
    function ProgressLinearInActive() {
        $("#StartProgressLinear").hide(function () {
            setTimeout(function () {
                $scope.ddeterminateValue = 0;
                $scope.showProgressLinear = true;
                $("#ProgressBgDiv").hide();
                if (!$scope.$$phase) {
                    $scope.$apply();
                }
            }, 500);
        });
    };
    function loadToast(alertMessage) {
        var el = document.querySelectorAll('#zoom');
        $mdToast.show(
            $mdToast.simple()
                .textContent(alertMessage)
                .position('bottom')
                .hideDelay(4000))
            .then(function () {
                $log.log('Toast dismissed.');
            }).catch(function () {
                $log.log('Toast failed or was forced to close early by another toast.');
            });
        if (!$scope.$$phase) {
            $scope.$apply();
        }
    };

    if (!$scope.$$phase) {
        $scope.$apply();
    }


})

app.controller('loginctrl', function ($scope, $mdDialog, $mdToast, $log, $location,) {


    $scope.login = function (ev) {
        ProgressLinearActive();
        var userObject = {

            email: $scope.useremail,
            password: $scope.password
            // Other properties of the user object
        };

        console.log(userObject)

        var settings = {
            "url": baseURL+ "/api/Home/Login",
            "method": "POST",
            "timeout": 0,
            "headers": {
                "Content-Type": "application/json"
            },
            "data": JSON.stringify({
                email: $scope.useremail,
                password: $scope.password
            }),
        };

        $.ajax(settings).done(function (response) {
            console.log(response);
            window.localStorage.setItem("userInfo", JSON.stringify(response))
            loadToast("Login Successful")
            $location.path("/")
            window.location.reload();

            ProgressLinearInActive()


        }).fail(function (error) {
            ProgressLinearInActive()

            console.log(error)
            loadToast(error.responseJSON.title)

        });


    }
    $scope.GotoSingup = function () {


        $location.path("/Signup")
    }
   

    function ProgressLinearActive() {
        $("#StartProgressLinear").show(function () {

            $("#ProgressBgDiv").show();
            $scope.ddeterminateValue = 15;
            $scope.showProgressLinear = false;
            if (!$scope.$$phase) {
                $scope.$apply();
            }
        });
    };
    function ProgressLinearInActive() {
        $("#StartProgressLinear").hide(function () {
            setTimeout(function () {
                $scope.ddeterminateValue = 0;
                $scope.showProgressLinear = true;
                $("#ProgressBgDiv").hide();
                if (!$scope.$$phase) {
                    $scope.$apply();
                }
            }, 500);
        });
    };
    function loadToast(alertMessage) {
        var el = document.querySelectorAll('#zoom');
        $mdToast.show(
            $mdToast.simple()
                .textContent(alertMessage)
                .position('bottom')
                .hideDelay(4000))
            .then(function () {
                $log.log('Toast dismissed.');
            }).catch(function () {
                $log.log('Toast failed or was forced to close early by another toast.');
            });
        if (!$scope.$$phase) {
            $scope.$apply();
        }
    };

    if (!$scope.$$phase) {
        $scope.$apply();
    }


})
app.controller('mainpagectrl', function ($scope, $mdDialog, $mdToast, $log, $http, $q, $location,) {
     
    var selcteddealInfo;
    $scope.MainContent = true;
    $scope.caluseList = true;
    $scope.LoadClauseBTN = true;


    let userinfo = window.localStorage.getItem("userInfo")
    userinfo = JSON.parse(userinfo)


    var userid;
    if (userinfo) {
        userid = userinfo.id;
       
        

    } else {
     

    }

    Office.onReady(function () {
      
    

    // Function to save the selected deal to localStorage
        let dialog; // Declare dialog as global for use in later functions.


        $scope.addCompany = function () {



            Office.context.ui.displayDialogAsync('https://localhost:44312/Templates/DealChatWebPage.html', { height: 80, width: 80 },
                function (asyncResult) {
                    dialog = asyncResult.value;
                    dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
                }
            );
        }

        function processMessage(arg) {
            dialog.close();
            const messageFromDialog = JSON.parse(arg.message);
            showUserName(messageFromDialog.name);
        }
    $scope.saveSelectedDeal = function (info) {
        localStorage.setItem('selectedDeal', JSON.stringify(info));
    };

    // Function to save the selected contract to localStorage
    $scope.saveSelectedContract = function (ClauseOfContracts) {
        localStorage.setItem('selectedContract', JSON.stringify(ClauseOfContracts));
    };
    $scope.refreshAddin = function () {

        window.location.reload();
    }
        Word.run(function (context) {
            var properties = context.document.properties;
            var customPropertyDeal = properties.customProperties.getItemOrNullObject("SelectedDeal");
            var customPropertyContract = properties.customProperties.getItemOrNullObject("SelectedContract");
            context.load(customPropertyDeal);
            context.load(customPropertyContract);

            return context.sync().then(function () {
                if (customPropertyDeal.isNullObject) {
                    //console.log("Custom property 'SelectedDeal' not found.");
                } else {
                    //console.log("Custom Property 'SelectedDeal' Value: " + customPropertyDeal.value);
                    let dealvalue = JSON.parse(customPropertyDeal.value);
                    $scope.Selected = dealvalue;
                    if (!$scope.$$phase) {
                        $scope.$apply();
                    }
                    //console.log("Parsed 'SelectedDeal' Value: ", dealvalue);
                }

                if (customPropertyContract.isNullObject) {
                    $scope.MainContent = true;

                    //if the custom property has null value 
                    //$scope.addContract();

                    //console.log("Custom property 'SelectedContract' not found.");
                } else {
                    //console.log("Custom Property 'SelectedContract' Value: " + customPropertyContract.value);
                    let contractvalue = JSON.parse(customPropertyContract.value);
                    if (contractvalue) {
                        $scope.MainContent = false;

                        $scope.GetClauses = contractvalue;
                        $.ajax({
                            type: "get",
                            url: baseURL+"/api/Home/getClause/" + contractvalue.id, // The URL of your controller action
                            //data: JSON.stringify(newItem),
                            contentType: "application/json; charset=utf-8",
                            dataType: "json",
                            success: function (response) {
                                console.log(response)

                                if (response.length > 0) {
                                    $scope.Clause = response
                                    $scope.caluseList = false;
                                    $scope.LoadClauseBTN = false;

                                } else {
                                    $scope.caluseList = true;
                                    $scope.LoadClauseBTN = true;


                                }




                                ProgressLinearInActive()

                                // console.log($scope.Companies)
                            },
                            error: function (error) {
                                ProgressLinearInActive()
                                // Handle error, e.g., display error message
                                console.error("Error adding item:", error);
                            }
                        });
                        console.log($scope.GetClauses);
                         
                        if (!$scope.$$phase) {
                            $scope.$apply();
                        }
                    } else { }

                  
                    //console.log("Parsed 'SelectedContract' Value: ", contractvalue);
                }
            });
        }).catch(function (error) {
            console.log(error);
        });
    $scope.getSelectedDeal = function () {
        // Find the selected deal object based on SelectedDeals value
        let info = $scope.Selected;
        Word.run(function (context) {
            var properties = context.document.properties;
            properties.customProperties.add("SelectedDeal", JSON.stringify(info));
            if (!$scope.$$phase) {
                $scope.$apply();
            }
            return context.sync();
        }).catch(function (error) {
            console.log(error);
        });

        //$scope.saveSelectedDeal(info)
        //selcteddealInfo = JSON.parse(info);
        if (info.dealList >= 0) {
            $scope.clauseTags = [];
            loadToast("This deal has no contracts")
        }
            $scope.clauseTags = [];

        $scope.dealContracts = info.dealList
        console.log(info);

        };


    var loadContractin_body = []
        var Selected_Contractid;
        $scope.GetContractClauses = function () {

        ProgressLinearActive();
            let ClauseOfContracts = $scope.GetClauses;
            Word.run(function (context) {
                var properties = context.document.properties;
                //properties.customProperties.add("SelectedContract", JSON.stringify(ClauseOfContracts));

                localStorage.setItem('selectedContractforclause', JSON.stringify(ClauseOfContracts));
                if (!$scope.$$phase) {
                    $scope.$apply();
                    }
                return context.sync();
            }).catch(function (error) {
                console.log(error);
            });
            //$scope.saveSelectedContract(ClauseOfContracts)
        //ClauseOfContracts = JSON.parse(ClauseOfContracts)
            loadContractin_body[0] = ClauseOfContracts
            Selected_Contractid = ClauseOfContracts.id
        console.log(ClauseOfContracts)


        //  call api to gte the clauses
        $.ajax({
            type: "get",
            url: baseURL+"/api/Home/getClause/" + ClauseOfContracts.id, // The URL of your controller action
            //data: JSON.stringify(newItem),
            contentType: "application/json; charset=utf-8",
            dataType: "json",
            success: function (response) {
                console.log(response)
                $scope.Clause = response
                $scope.caluseList = false;




                ProgressLinearInActive()

                // console.log($scope.Companies)
            },
            error: function (error) {
                ProgressLinearInActive()
                // Handle error, e.g., display error message
                console.error("Error adding item:", error);
            }
        });

        // $scope.SelectedUserEmail = userEmail.email


        }

        $scope.getselectedClause = function (clause) {

            // Reset all items to unselected
            $scope.Clause.forEach(function (item) {
                item.selected = false;
            });

            // Set the clicked item as selected
            clause.selected = true;
            ProgressLinearActive()
            console.log(clause)
            $.ajax({
                type: "get",
                url: baseURL+"/api/Home/getTagClause?clauseid=" + clause.id, // The URL of your controller action
                //data: JSON.stringify(newItem),
                contentType: "application/json; charset=utf-8",
                dataType: "json",
                success: function (response) {
                    console.log(response)
                    if (response.length > 0) {

                        $scope.hidetitle = false;


                    }
                    $scope.clauseTags = response;

                    /// autoselect the text of clause name from word body
                    Word.run(function (context) {
                        var body = context.document.body;
                        var searchResults = body.search(clause.name, { matchCase: false });
                        context.load(searchResults);
                        return context.sync().then(function () {
                            if (searchResults.items.length > 0) {
                                var firstResult = searchResults.items[0];
                                firstResult.select();

                                //$.ajax({
                                //    type: "get",
                                //    url: baseURL+"/api/Home/getTagClause?clauseid=" + clause.id, // The URL of your controller action
                                //    //data: JSON.stringify(newItem),
                                //    contentType: "application/json; charset=utf-8",
                                //    dataType: "json",
                                //    success: function (response) {
                                //        console.log(response)
                                //        if (response.length > 0) {

                                //            $scope.hidetitle = false;


                                //        } 
                                //        $scope.clauseTags = response;

                                //        $(document).ready(function () {
                                //            const scrollingElement = (document.scrollingElement || document.body);

                                //            const scrollSmoothToBottom = () => {
                                //                $(scrollingElement).animate({
                                //                    scrollTop: document.body.scrollHeight,
                                //                }, 500);
                                //            }
                                //            scrollSmoothToBottom();





                                //        })


                                //        ProgressLinearInActive()

                                //        // console.log($scope.Companies)
                                //    },
                                //    error: function (error) {
                                //        ProgressLinearInActive();

                                //        $scope.clauseTags = []
                                //        if (!$scope.$$phase) {
                                //            $scope.$apply();
                                //        }
                                //        if (error.status===404) {
                                //            loadToast("No Tagged Email Found")

                                //        }
                                //        // Handle error, e.g., display error message
                                //        console.error("Error adding item:", error.status);
                                //    }
                                //})

                                console.log("Text found and selected: " + firstResult.text);
                            } else {
                                console.log("Text not found.");
                                ProgressLinearInActive();

                            }
                        });
                    }).catch(function (error) {
                        console.error("Error: " + JSON.stringify(error));
                        ProgressLinearInActive();

                    });





                    ///scrol down the addin screen after loading response
                    $(document).ready(function () {
                        const scrollingElement = (document.scrollingElement || document.body);

                        const scrollSmoothToBottom = () => {
                            $(scrollingElement).animate({
                                scrollTop: document.body.scrollHeight,
                            }, 500);
                        }
                        scrollSmoothToBottom();

                    })




                    ProgressLinearInActive()

                    // console.log($scope.Companies)
                },
                error: function (error) {
                    ProgressLinearInActive();

                    $scope.clauseTags = []
                    if (!$scope.$$phase) {
                        $scope.$apply();
                    }
                    if (error.status === 404) {
                        loadToast("No Tagged Email Found")

                    }
                    // Handle error, e.g., display error message
                    console.error("Error adding item:", error.status);
                }
            })
               

        }

        $scope.PostNewClause = function (ev) {

            ProgressLinearActive();
            
            $mdDialog.show({

                scope: $scope.$new(),
                templateUrl: 'Templates/addClause.html',
                targetEvent: ev,
                fullscreen: $scope.customFullscreen,
                controller: ['$scope', '$mdDialog', function ($scope, $mdDialog) {
                    //new  ode
                    $scope.checkedItems = []; // Array to store checked items

                    $scope.boldLines = [];

                    Word.run(function (context) {
                        var body = context.document.body;
                        var paragraphs = body.paragraphs;

                        context.load(paragraphs, 'text, font');
                        return context.sync()
                            .then(function () {
                                //var boldLines = [];

                                // Loop through the paragraphs looking for bolded lines
                                paragraphs.items.forEach(function (paragraph) {
                                    if (paragraph.font.bold) {
                                        $scope.boldLines.push(paragraph.text);
                                        ProgressLinearInActive();

                                    }
                                });



                                $scope.filteredBoldLines = $scope.cleanArray($scope.boldLines);

                                // Log the array of bold lines
                                console.log("Bold Lines:", $scope.boldLines);
                            })
                            .catch(function (error) {
                                ProgressLinearInActive();
                                console.log(error);
                            });
                    });

                    $scope.cleanArray = function (arr) {
                        // Remove empty strings and duplicates
                        return arr.filter(function (item, index) {
                            return item.trim() !== '' && arr.indexOf(item) === index;
                        });
                    };
                    //$scope.selected = $scope.boldLines;
                    //$scope.items = [1, 2, 3, 4, 5];
                    $scope.selected = [];
                    $scope.toggle = function (item, list) {
                        var idx = list.indexOf(item);
                        if (idx > -1) {
                            list.splice(idx, 1);
                        }
                        else {
                            list.push(item);
                        }
                    };

                    $scope.exists = function (item, list) {
                        return list.indexOf(item) > -1;
                    };

                    $scope.isIndeterminate = function () {
                        return ($scope.selected.length !== 0 &&
                            $scope.selected.length !== $scope.boldLines.length);
                    };

                    $scope.isChecked = function () {
                        return $scope.selected.length === $scope.boldLines.length;
                    };

                    $scope.toggleAll = function () {
                        if ($scope.selected.length === $scope.boldLines.length) {
                            $scope.selected = [];
                            console.log($scope.selected)
                        } else if ($scope.selected.length === 0 || $scope.selected.length > 0) {
                            $scope.selected = $scope.boldLines.slice(0);
                            console.log($scope.selected)

                        }
                    };
                    //$scope.items = [1, 2, 3, 4, 5];
                    //$scope.selected = [];

                    //$scope.toggle = function (item, list) {
                    //    var idx = list.indexOf(item);
                    //    if (idx > -1) {
                    //        list.splice(idx, 1);
                    //    }
                    //    else {
                    //        list.push(item);
                    //    }
                    //};

                    //$scope.exists = function (item, list) {
                    //    return list.indexOf(item) > -1;
                    //};
                    let selected_Contract = localStorage.getItem('selectedContractforclause');
                    selected_Contract = JSON.parse(selected_Contract);
                    console.log(selected_Contract)

                    //new  ode end
                    //old code

                    //$scope.ClauseTitle = "";
                    //$scope.ClauseDetail = "";
                    
                    //Office.context.document.getSelectedDataAsync(Office.CoercionType.Text, function (asyncResult) {
                    //    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
                    //        //write('Action failed. Error: ' + asyncResult.error.message);
                    //    }
                    //    else {

                    //        Word.run(function (context) {
                    //            var words = context.document.getSelection().getTextRanges([' '], true);
                    //            console.log(words)
                    //            context.load(words, ['text', 'font']);
                    //            var boldRanges = [];
                    //            return context.sync()
                    //                .then(function () {

                    //                    console.log(words.text)
                    //                    for (var i = 0; i < words.items.length; ++i) {
                    //                        var word = words.items[i];
                    //                        if (word.font.bold)
                    //                            boldRanges.push(word);
                    //                    }
                    //                })
                    //                .then(function () {

                    //                    console.log(boldRanges)
                    //                    console.log(boldRanges[0]._Te)
                    //                    let valueSelected = asyncResult.value;
                    //                    let ValueBold = boldRanges[0]._Te;

                    //                    var originalString = valueSelected;
                    //                    var startIndex = originalString.indexOf(ValueBold);
                    //                    var endIndex = startIndex + ValueBold.length;
                    //                    var modifiedString = originalString.substring(0, startIndex) + originalString.substring(endIndex);
                    //                    console.log(modifiedString);

                    //                    $scope.ClauseTitle = ValueBold;
                    //                    $scope.ClauseDetail = modifiedString;


                    //                    //var neworiginalString = "/r";
                    //                    //var startIndexes = neworiginalString.indexOf($scope.ClauseDetail);
                    //                    //var endIndexes = startIndexes + $scope.ClauseDetail.length;
                    //                    //var modifiedStringddd = neworiginalString.substring(0, startIndexes) + neworiginalString.substring(endIndexes);
                    //                    //console.log(modifiedStringddd);
                    //                    console.log($scope.ClauseDetail)
                    //                    //for (var j = 0; j < boldRanges.length; ++j) {
                    //                    //    boldRanges[j].font.highlightColor = '#FF00FF';
                    //                    //}
                    //                });
                    //        });

                    //    }
                    //});





                    ////load the contract clauses from json file or database


                    ////////////////////////////////////////////////

                    $scope.Save = function () {
                        var promises = [];
                        if ($scope.selected.length > 0) {
                            ProgressLinearActive();
                         

                            for (let i = 0; i < $scope.selected.length; i++) {
                                var promise = $http({
                                    method: "POST",
                                    url: baseURL+"/api/Home/NewClause",
                                    data: JSON.stringify({
                                        "name": $scope.selected[i],
                                        "contractid": $scope.GetClauses.id
                                    }),
                                    headers: {
                                        "Content-Type": "application/json; charset=utf-8"
                                    }
                                });

                                promises.push(promise);
                            }

                            $q.all(promises).then(function (results) {
                                // All requests have completed successfully
                                console.log("All clauses added successfully:", results);

                                ProgressLinearInActive();
                                window.location.reload();
                                loadToast("All clauses added successfully");
                            }).catch(function (error) {
                                // Handle errors
                                ProgressLinearInActive();

                                console.error("Error adding clauses:", error);
                                loadToast("Error adding clauses");
                            });
                        } else {

                            loadToast("Please select clause to save")
                            //ProgressLinearInActive();

                        }


                    }





                    $scope.closeDialog = function () {
                        $mdDialog.hide();
                    };





                }]

            })







         
        }
      
        $scope.loadContracts = function () {

           

            Word.run(function (context) {
                // Get the active document
                var document = context.document;

                // Get the body of the document
                var body = document.body;

                // Loop through the data array and insert each item
                for (var i = 0; i < loadContractin_body.length; i++) {
                    var item = loadContractin_body[i];

                    // Insert the key with bold formatting
                    var keyParagraph = body.insertParagraph(item.contractName + ": ", Word.InsertLocation.end);
                    keyParagraph.font.bold = true;

                    // Insert the value without bold formatting
                    var valueParagraph = body.insertParagraph(item.contractDetail, Word.InsertLocation.end);
                    valueParagraph.font.bold = false
                    // Add a new line between items
                    body.insertParagraph("", Word.InsertLocation.end);
                }

                // Synchronize the document changes
                return context.sync();
            }).catch(function (error) {
                console.log(error);
            });


        }
        
        $scope.hidetitle = true;

        $scope.LoadClause = function () {


            ProgressLinearActive();
            //<-------    laod all clauses code starts here   ----------->

            //for (i = 0; i < $scope.Clause.length; i++) {

            //$.ajax({
            //    type: "get",
            //    url: baseURL+"/api/Home/getTagClause?clauseid=" + $scope.Clause[i].id, // The URL of your controller action
            //    //data: JSON.stringify(newItem),
            //    contentType: "application/json; charset=utf-8",
            //    dataType: "json",
            //    success: function (response) {
            //        console.log(response)
            //        $scope.clauseTags = response

            //        for (i = 0; i < $scope.Clause.length; i++) {
            //            if ($scope.Clause[i].id === response[i].clauseid) {

            //                $scope.titlename = $scope.Clause[i].name;
            //                console.log($scope.titlename)

            //            }

            //        }


            //        //if (response.length > 0) {

            //        //    $scope.hidetitle = false;


            //        //}



            //        ProgressLinearInActive()

            //        // console.log($scope.Companies)
            //    },
            //    error: function (error) {
            //        ProgressLinearInActive()

            //        $scope.clauseTags = []
            //        if (!$scope.$$phase) {
            //            $scope.$apply();
            //        }

            //        // Handle error, e.g., display error message
            //        console.error("Error adding item:", error);
            //    }
            //})
            //}

            //<-------    laod all clauses code ends here   ----------->





            Word.run(function (context) {
                var selection = context.document.getSelection();
                selection.load("text");

                return context.sync().then(function () {
                    var selectedText = selection.text.trim(); // Trim whitespace
                    console.log($scope.Clause);
                    var matchedClauses = $scope.Clause.filter(cls => cls.name === selectedText);

                    if (matchedClauses.length > 0) {
                        console.log("Matches found:");
                        console.log(matchedClauses);
                        $scope.ClasueTitleName = matchedClauses[0].name
                        $scope.getselectedClause(matchedClauses[0]);
                        //$.ajax({
                        //    type: "get",
                        //    url: baseURL+"/api/Home/getTagClause?clauseid=" + matchedClauses[0].id, // The URL of your controller action
                        //    //data: JSON.stringify(newItem),
                        //    contentType: "application/json; charset=utf-8",
                        //    dataType: "json",
                        //    success: function (response) {
                        //        console.log(response)
                        //        if (response.length>0) {
                        //            $(document).ready(function () {
                        //                const scrollingElement = (document.scrollingElement || document.body);

                        //                const scrollSmoothToBottom = () => {
                        //                    $(scrollingElement).animate({
                        //                        scrollTop: document.body.scrollHeight,
                        //                    }, 500);
                        //                }
                        //                scrollSmoothToBottom();

                        //            })
                        //            $scope.hidetitle = false;


                        //        }
                        //        $scope.clauseTags = response




                        //        ProgressLinearInActive()

                        //        // console.log($scope.Companies)
                        //    },
                        //    error: function (error) {
                        //        ProgressLinearInActive()

                        //        $scope.clauseTags = []
                        //        if (!$scope.$$phase) {
                        //            $scope.$apply();
                        //        }

                        //        // Handle error, e.g., display error message
                        //        console.error("Error adding item:", error);
                        //    }
                        //})
                        //$scope.foundContracts.push(...matchedClauses);
                        ProgressLinearInActive();
                    } else {
                        ProgressLinearInActive();

                        loadToast("No matching Clauses found for the selection", selectedText);
                    }
                });

            }).catch(function (error) {
                console.log(error.message);
                ProgressLinearInActive();

            });
        };


        $scope.ViewMialBody = function (items) {
            //console.log(html)
            console.log(items.htmlContent)
            let html = items.htmlContent
            $scope.showiconhtml = items.htmlContent
          
                var encodedHtmlContent = encodeURIComponent(html);
            // Function to show a dialog box
            window.localStorage.setItem('htmlContent', JSON.stringify(html));

         

                Office.context.ui.displayDialogAsync(
                    'https://localhost:44312/Templates/MailHtmlBody.html',
                    { height: 50, width: 50 },
                    function (result) {
                        var dialog = result.value;

                        // Add event handlers for dialog box
                        dialog.addEventHandler(Office.EventType.DialogMessageReceived, function (args) {
                            var messageFromDialog = args.getArgs();
                            console.log('Message from dialog: ' + messageFromDialog);

                            // Close the dialog box if needed
                            dialog.close();
                        });

                        // Handle dialog box closed event
                        dialog.addEventHandler(Office.EventType.DialogEventReceived, function (args) {
                            var dialogClosedReason = args.getEventArgs().getReason();
                            if (dialogClosedReason === Office.MailboxEnums.DialogEventReasons.DialogClosed) {
                                console.log('Dialog box closed by user.');
                            }
                        });
                    }
                );

           

            // Call the showDialog function to open the dialog box

        }
      
        $scope.addCompany = function () {



            Office.context.ui.displayDialogAsync('https://localhost:44312/Templates/DealChatWebPage.html?userId=' + userid, { height: 80, width: 80 },
                function (asyncResult) {
                    dialog = asyncResult.value;
                    dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
                }
            );
        }

    })

    function ProgressLinearActive() {
        $("#StartProgressLinear").show(function () {

            $("#ProgressBgDiv").show();
            $scope.ddeterminateValue = 15;
            $scope.showProgressLinear = false;
            if (!$scope.$$phase) {
                $scope.$apply();
            }
        });
    };
    function ProgressLinearInActive() {
        $("#StartProgressLinear").hide(function () {
            setTimeout(function () {
                $scope.ddeterminateValue = 0;
                $scope.showProgressLinear = true;
                $("#ProgressBgDiv").hide();
                if (!$scope.$$phase) {
                    $scope.$apply();
                }
            }, 500);
        });
    };
    function loadToast(alertMessage) {
        var el = document.querySelectorAll('#zoom');
        $mdToast.show(
            $mdToast.simple()
                .textContent(alertMessage)
                .position('bottom')
                .hideDelay(4000))
            .then(function () {
                $log.log('Toast dismissed.');
            }).catch(function () {
                $log.log('Toast failed or was forced to close early by another toast.');
            });
        if (!$scope.$$phase) {
            $scope.$apply();
        }
    };

    if (!$scope.$$phase) {
        $scope.$apply();
    }

})
app.controller('DealChatCTRL', function ($scope, $mdDialog, $mdToast, $log, $location,) {
    $scope.Contracts = [];
    var userlog = false;
    //var baseURL = "https://localhost:44371/api"
    $scope.password = "";
    $scope.useremail = "";

   //let userinfo = window.localStorage.getItem("userInfo")
   // userinfo = json.parse(userinfo)
    //if (userlog == false) {


    //    $location.path("/Login")
    //}








    let userinfo = window.localStorage.getItem("userInfo")
    userinfo = JSON.parse(userinfo)


    var userid;
    if (userinfo) {
        userid = userinfo.id;
        console.log("userid :>"+userid)
        if (userinfo.id) {
            $location.path("/")
        }

    } else {
        $location.path("/Login")


    }


    //get Deals 
  

    Office.onReady(function () {

        ProgressLinearActive()
        getuseDeals();

        function getuseDeals() {
            $.ajax({
                type: "get",
                url: baseURL+"/api/Home/GetDeal/" + userid, // The URL of your controller action
                //data: JSON.stringify(newItem),
                contentType: "application/json; charset=utf-8",
                dataType: "json",
                success: function (response) {
                    console.log(response)
                    $scope.Deals = response
                    ProgressLinearInActive()

                    console.log($scope.Deals)
                },
                error: function (error) {
                    ProgressLinearInActive()
                    // Handle error, e.g., display error message
                    console.error("Error adding item:", error);
                }
            });

        }
      
        var selcteddealInfo;


     
      

   
   

  

   

        $scope.show_Tag_sEmail = function (objs) {

            console.log(objs)
        }

        //<--- add new contract process----->
        
        $scope.addContract = function (ev) {


            $mdDialog.show({

                scope: $scope.$new(),
                templateUrl: 'Templates/addContract.html',
                targetEvent: ev,
                fullscreen: $scope.customFullscreen,
                controller: ['$scope', '$mdDialog', function ($scope, $mdDialog) {
                    var selcteddealInfo;

                    $.ajax({
                        type: "get",
                        url: baseURL+ "/api/Home/GetDeal/" + userid, // The URL of your controller action
                        //data: JSON.stringify(newItem),
                        contentType: "application/json; charset=utf-8",
                        dataType: "json",
                        success: function (response) {
                            console.log(response)
                            $scope.Deals = response
                            ProgressLinearInActive()

                        },
                        error: function (error) {
                            ProgressLinearInActive()
                            // Handle error, e.g., display error message
                            console.error("Error adding item:", error);
                        }
                    });




                    $scope.ClauseTitle = "";
                    $scope.ClauseDescription = "";


                    Word.run(function (context) {
                        // Insert your code here. For example:
                        var documentBody = context.document.body;
                        context.load(documentBody);
                        return context.sync()
                            .then(function () {

                                let wordBodyText = documentBody.text;

                                console.log(documentBody.text);

                                if (wordBodyText) {
                                    //$scope.ContractTitle = selectedText
                                    $scope.ContractDescription = wordBodyText
                                    ProgressLinearInActive();
                                } else {
                                    ProgressLinearInActive();

                                }

                            })
                    });
                
                    ////load the contract clauses from json file or database
              

                    ////////////////////////////////////////////////
                    var DealCustomProperty = {};
                    $scope.getSelectedDeals = function () {
                        // Find the selected deal object based on SelectedDeals value
                        let info = $scope.SelectedDeals;
                        selcteddealInfo = JSON.parse(info);
                        DealCustomProperty = selcteddealInfo;
                        console.log(selcteddealInfo);

                    };
                    $scope.Save = function () {
                        ProgressLinearActive();
                     
                        $.ajax({
                            type: "POST",
                            url: baseURL+"/api/Home/NewContract", // The URL of your controller action
                            data: JSON.stringify({
                                "id": "3fa85f64-5717-4562-b3fc-2c963f66afa6",
                                "contractName": $scope.ContractTitle,
                                "contractDetail": $scope.ContractDescription,
                                "dealid": selcteddealInfo.id,
                                "userid": userid
                            }),
                            contentType: "application/json; charset=utf-8",
                            //dataType: "json",
                            success: function (data) {
                                // Handle success, e.g., update UI, display message, etc.
                                console.log("Item added succeGetContractClausesssfully:", data);
                                $scope.GetClauses = data;
                                Word.run(function (context) {
                                    var properties = context.document.properties;
                                    properties.customProperties.add("SelectedDeal", JSON.stringify(DealCustomProperty));

                                    properties.customProperties.add("SelectedContract", JSON.stringify(data));

                                    localStorage.setItem('selectedContractforclause', JSON.stringify(data));

                                    if (!$scope.$$phase) {
                                        $scope.$apply();
                                    }


                                    $mdDialog.hide();
                                    //getdefaultContracts()

                                    ProgressLinearInActive();
                                    window.location.reload();

                                    loadToast("Contract added successfuly")
                                    return context.sync();
                                }).catch(function (error) {
                                    console.log(error);
                                    ProgressLinearInActive()
                                });

                            },
                            error: function (error) {
                                // Handle error, e.g., display error message
                                console.error("Error adding item:", error);
                                loadToast("Error adding Contract")

                                $mdDialog.hide();


                            }
                        });
            
                    };




                    $scope.closeDialog = function () {
                        $mdDialog.hide();
                    };

                



                }]

            })

        }

        var selectedtext = "";
        //$scope.foundContracts = [];
        //$scope.foundContracts;
     
        $scope.GotoSingup = function () {


            $location.path("/Signup")
        }
        $scope.GotoSingIn = function () {


            $location.path("/Login")
        }
        $scope.Logout = function () {
            ProgressLinearActive();

            window.localStorage.clear("userInfo")
            $location.path("/Login")
            ProgressLinearInActive();
        }

        
        function ProgressLinearActive() {
            $("#StartProgressLinear").show(function () {

                $("#ProgressBgDiv").show();
                $scope.ddeterminateValue = 15;
                $scope.showProgressLinear = false;
                if (!$scope.$$phase) {
                    $scope.$apply();
                }
            });
        };
        function ProgressLinearInActive() {
            $("#StartProgressLinear").hide(function () {
                setTimeout(function () {
                    $scope.ddeterminateValue = 0;
                    $scope.showProgressLinear = true;
                    $("#ProgressBgDiv").hide();
                    if (!$scope.$$phase) {
                        $scope.$apply();
                    }
                }, 500);
            });
        };
        function loadToast(alertMessage) {
            var el = document.querySelectorAll('#zoom');
            $mdToast.show(
                $mdToast.simple()
                    .textContent(alertMessage)
                    .position('bottom')
                    .hideDelay(4000))
                .then(function () {
                    $log.log('Toast dismissed.');
                }).catch(function () {
                    $log.log('Toast failed or was forced to close early by another toast.');
                });
            if (!$scope.$$phase) {
                $scope.$apply();
            }
        };

        if (!$scope.$$phase) {
            $scope.$apply();
        }


    });
})
