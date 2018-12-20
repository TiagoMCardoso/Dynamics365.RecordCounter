if (typeof (RecordCounter) == "undefined") { RecordCounter = { __namespace: true }; }

RecordCounter.APIVersion = "v9.1";
RecordCounter.Counter = 5;
RecordCounter.EntityPluralNamesTemp = [];
RecordCounter.EntityPluralNames = [];

RecordCounter.OnLoad = function () {
    // Retrieve All Dynamics 365 Entities (Plural Name)
    RecordCounter.RetrieveEntities();

    // OnClick - Add New Row
    $("#addrow").on("click", function () {
        //debugger;
        var newRow = $("<tr>");
        var cols = "";

        cols += "<td><div class='ui-widget'><input class='form-control' id='entity" + RecordCounter.Counter + "' name='entity'></div></td>";
        cols += "<td><div class='form-group'><input class='form-control' type='text' id='record" + RecordCounter.Counter + "' name='total' placeholder='' readonly /</div></td>";
        cols += "<td><div id='loader" + RecordCounter.Counter + "' name='loader' class='loader' style='display: none'></div></td>";
        cols += '<td><input type="button" class="ibtnDel btn btn-md btn-danger" value="Delete Row" /></td>';

        newRow.append(cols);

        $("table.order-list").append(newRow);

        // Set AutoComplete (All Entities)
        RecordCounter.SetAutocomplete("#entity" + RecordCounter.Counter);

        RecordCounter.Counter++;
    });

    // OnClick - Delete Row
    $("table.order-list").on("click", ".ibtnDel", function (event) {
        $(this).closest("tr").remove();
        RecordCounter.Counter -= 1
    });
},

// Retrieve All Dynamics Entities
RecordCounter.RetrieveEntities = function () {
    try {
        var clientURL = Xrm.Page.context.getClientUrl();
        var req = new XMLHttpRequest();

        req.open("GET", encodeURI(clientURL + "/api/data/" + RecordCounter.APIVersion + "/"), true);
        req.setRequestHeader("Accept", "application/json");
        req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
        req.setRequestHeader("OData-MaxVersion", "4.0");
        req.setRequestHeader("OData-Version", "4.0");
        req.onreadystatechange = function () {
            if (this.readyState == 4 /* complete */) {
                req.onreadystatechange = null;
                if (this.status == 200)//to verify result is OK
                {
                    //debugger;
                    var data = JSON.parse(this.response);

                    if (data != null) {
                        RecordCounter.EntityPluralNamesTemp = data.value;

                        // Populate PluralName Array
                        RecordCounter.PopulateEntityPluralNameArray();

                        // Set AutoComplete (All Entities)
                        RecordCounter.SetAutocomplete("#entity0");
                        RecordCounter.SetAutocomplete("#entity1");
                        RecordCounter.SetAutocomplete("#entity2");
                        RecordCounter.SetAutocomplete("#entity3");
                        RecordCounter.SetAutocomplete("#entity4");
                    }
                }
                else {
                    var error = JSON.parse(this.response).error;
                    RecordCounter.Error(error.message, null);
                }
            }
        };
        req.send(null);
    } catch (error) {
        RecordCounter.Error(error, null);
    }
},

// Populate Entity Plural Name Array
RecordCounter.PopulateEntityPluralNameArray = function () {
    try {
        // Lopp Entities
        for (var i = 0; i < RecordCounter.EntityPluralNamesTemp.length; i++) {
            RecordCounter.EntityPluralNames.push(RecordCounter.EntityPluralNamesTemp[i].name);

            // *** CODE TO COUNT ALL DYNAMICS 365 ENTITIES ***
            //var newRow = $("<tr>");
            //var cols = "";

            //cols += "<td><div class='ui-widget'><input class='form-control' id='entity" + RecordCounter.Counter + "' name='entity' value='" + RecordCounter.EntityPluralNamesTemp[i].name + "'></div></td>";
            //cols += "<td><div class='form-group'><input class='form-control' type='text' id='record" + RecordCounter.Counter + "' name='total' placeholder='' readonly /</div></td>";
            //cols += "<td><div id='loader" + RecordCounter.Counter + "' name='loader' class='loader' style='display: none'></div></td>";
            //cols += '<td><input type="button" class="ibtnDel btn btn-md btn-danger" value="Delete Row" /></td>';

            //newRow.append(cols);

            //$("table.order-list").append(newRow);

            //// Set AutoComplete (All Entities)
            ////SetAutocomplete("#entity" + RecordCounter.Counter);

            //RecordCounter.Counter++;
            // *** END ***
        }
    } catch (error) {
        RecordCounter.Error(error, null);
    }
},

// Set AutoComplete (All Entities)
RecordCounter.SetAutocomplete = function (entityAutocompleteId) {
    try {
        $(entityAutocompleteId).autocomplete({
            source: RecordCounter.EntityPluralNames,
            delay: 300,
            minLength: 0,
        });
    } catch (error) {
        RecordCounter.Error(error, null);
    }
},

// Count Records of the mentioned entities
RecordCounter.ExecuteCount = function () {
    try {
        // Clean all entity counters
        RecordCounter.CleanTotal();

        //debugger;
        $('#myTable > tbody > tr').each(function () {

            var totalId = "#" + $(this).find('input[name="total"]').attr('id');
            var entityPluralName = $(this).find('input[name="entity"]').val();
            var loaderId = "#" + $(this).find('div[name="loader"]').attr('id');

            if (entityPluralName != "") {
                $(loaderId).show();

                // 
                RecordCounter.CountRecords(totalId, loaderId, entityPluralName, null, null);
            }
        });

        // Enable/Disable all the buttons
        RecordCounter.EnableDisableAllButtons(true);
    } catch (error) {
        RecordCounter.Error(error, null);
    }
},

RecordCounter.CountRecords = function (totalId, loaderId, entityPluralName, nextLink, lastError) {
    try {
        //debugger;
        var query;

        // Check if this entity doesn't have pagination (when entity is bigger then 5,000 records)
        if (nextLink == null) {
            var clientURL = Xrm.Page.context.getClientUrl();

            // Check if the last try/execution wasn't return any error
            if (lastError == null) {
                query = encodeURI(clientURL + "/api/data/" + RecordCounter.APIVersion + "/" + entityPluralName + "?$select=createdon&$count=true");
            }
            // If last try/execution has returned an error, try to count without a "select" filter
            else {
                query = encodeURI(clientURL + "/api/data/" + RecordCounter.APIVersion + "/" + entityPluralName + "?$count=true");
            }
        }
        // If entity has pagination, simply call the next page
        else {
            query = nextLink;
        }

        var req = new XMLHttpRequest();
        req.open("GET", query, true);
        req.setRequestHeader("Accept", "application/json");
        req.setRequestHeader("Content-Type", "application/json; charset=utf-8");
        req.setRequestHeader("OData-MaxVersion", "4.0");
        req.setRequestHeader("OData-Version", "4.0");
        req.onreadystatechange = function () {
            if (this.readyState == 4 /* complete */) {
                req.onreadystatechange = null;
                if (this.status == 200)//to verify result is OK
                {
                    var data = JSON.parse(this.response);

                    //debugger;
                    // Check if call has returned a valid result
                    if (data != null && data["@odata.error"] == undefined) {

                        var value;
                        var nextLink;

                        // Remove Readonly class for all Total fields
                        $(totalId).removeAttr('readonly');

                        // Check if Total still NULL/BLANK
                        if (($(totalId).val()) == "") {
                            value = 0;
                        }
                        // If Total already has value remove the format
                        else {
                            value = $(totalId).val().replace(",", "");
                            value = parseInt(value);
                        }

                        // Add new results into the counter
                        value += data.value.length;
                        value = value.toString();
                        value = RecordCounter.AddCommas(value);

                        // Set Total's value
                        $(totalId).val(value);

                        // Add Redonly class back for all Total fields
                        $(totalId).prop('readonly', true);

                        // Get the next page in the case of entity has more than 5,000 records
                        nextLink = data['@odata.nextLink'];

                        // Check if there a next page to be called
                        if (nextLink != null) {
                            // Call CounterRecords function again until all entity records being counted
                            RecordCounter.CountRecords(totalId, loaderId, null, nextLink, null);
                        }
                        else {
                            // If there is no more pages hide the loader
                            $(loaderId).hide();
                        }
                    }
                    else {
                        var error = data['@odata.error'];
                        RecordCounter.Error(error.message, loaderId);
                    }
                }
                else {
                    var error = JSON.parse(this.response).error;

                    // Check if error message refers to "Could not find a property named 'createdon'"
                    // If yes, call CountRecords function again, but trying to get counter from API without specified the "select" filter
                    if (error.message.indexOf("Could not find a property named 'createdon'") >= 0) {
                        RecordCounter.CountRecords(totalId, loaderId, entityPluralName, nextLink, error);
                    }
                    else {
                        // Check if error message refers to "Resource not found for the segment"
                        // If yes, add more information into the error message
                        if (error.message.indexOf("Resource not found for the segment") >= 0) {
                            RecordCounter.Error("There is no Entity with Plural Name called <b>" + entityPluralName + "</b>. Please check the spelling and confirm that you are using the Entity Plural Name (I.e.: account<b>S</b>, contact<b>S</b>, opportunit<b>IES</b>).", loaderId);
                        }
                        else {
                            RecordCounter.Error(error.message, loaderId);
                        }
                    }
                }
            }
        };
        req.send(null);
    } catch (error) {
        RecordCounter.Error(error, null);
    }
},

// Enable or Disable all Buttons in the form
RecordCounter.EnableDisableAllButtons = function (enabled) {
    try {
        if (enabled) {
            // Keep checking until the last entity counter finishes
            var time = setInterval(function () {
                if (!$(".loader").is(":visible")) {
                    // Enable all buttons
                    $("#btnExecute").removeAttr("disabled");
                    $(".ibtnNew").removeAttr("disabled");
                    $(".ibtnDel").removeAttr("disabled");

                    clearInterval(time);
                }
            },
            500);
        }
        else {
            // Disable all buttons
            $("#btnExecute").attr("disabled", true);
            $('.ibtnNew').prop('disabled', true);
            $('.ibtnDel').prop('disabled', true);
        }
    } catch (error) {
        RecordCounter.Error(error, null);
    }
},

RecordCounter.CleanTotal = function () {
    try {
        // Disable all buttons
        RecordCounter.EnableDisableAllButtons(false);

        // Clean error message box
        $("#errorBox").hide();

        // Clean all Total field values
        $('#myTable > tbody > tr').each(function () {
            $(this).find('input[name="total"]').removeAttr('readonly').val(null);
            $(this).find('input[name="total"]').prop('readonly', true);
        });
    } catch (error) {
        RecordCounter.Error(error, null);
    }
},

RecordCounter.Error = function (message, loaderId) {
    try {
        // Enable all buttons
        RecordCounter.EnableDisableAllButtons(true);

        // Show error messsage box
        $("#errorBox").show();

        // Append new error message into the current value
        $("#errorMessage").html($("#errorMessage").html() + message + "<br/>");

        if (loaderId != null) {
            $(loaderId).hide();
        }
    } catch (error) {
        RecordCounter.Error(error, null);
    }
},

RecordCounter.AddCommas = function (nStr) {
    try {
        nStr += '';
        x = nStr.split('.');
        x1 = x[0];
        x2 = x.length > 1 ? '.' + x[1] : '';
        var rgx = /(\d+)(\d{3})/;
        while (rgx.test(x1)) {
            x1 = x1.replace(rgx, '$1' + ',' + '$2');
        }
        return x1 + x2;
    }
    catch (error) {
        RecordCounter.Error(error, null);
    }
}