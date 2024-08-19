debugger;
// import { PublicClientApplication } from "@azure/msal-browser";
async function fetchDataAndSendToSdmAPI() {
  var formContext = Xrm.Page;
  await saveCurrentPage(formContext, false);
  var incidentId = formContext.data.entity.getId();
  var sdmBody = {
    referenceOrderData: {
      id: "",
      flexFields: [],
    },
  };

  var sdmSuccess = false;
  var incidentFetchXml =
    "<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>" +
    "<entity name='incident'>" +
    "<all-attributes />" +
    "<filter type='and'>" +
    "<condition attribute='incidentid' operator='eq' value='" +
    incidentId +
    "' />" +
    "</filter>" +
    "</entity>" +
    "</fetch>";

  try {
    const results = await Xrm.WebApi.retrieveMultipleRecords(
      "incident",
      `?fetchXml=${incidentFetchXml}`
    );
    if (results.entities.length > 0) {
      const incidentRecord = results.entities[0];
      sdmBody.referenceOrderData.id = incidentRecord["title"];
      var dcar_sectioncategory = incidentRecord["dcar_sectioncategory"];
      const servicestatus = formContext.getAttribute("dcar_servicestatus");
      const serviceevent = formContext.getAttribute("dcar_serviceevent");

      switch (dcar_sectioncategory) {
        case 1: //dcar_sectioncategory == collection
          //========================= repair details  ============================//
          var carRepairDetailsid =
            incidentRecord["_dcar_carrepairdetails_value"];
          var carRepairDetailsFxml =
            "<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>" +
            "<entity name='dcar_carrepairdetails'>" +
            "<all-attributes />" +
            "<filter type='and'>" +
            "<condition attribute='dcar_carrepairdetailsid' operator='eq' value='" +
            carRepairDetailsid +
            "' />" +
            "</filter>" +
            "</entity>" +
            "</fetch>";

          try {
            const carRepairDetailsResults =
              await Xrm.WebApi.retrieveMultipleRecords(
                "dcar_carrepairdetails",
                `?fetchXml=${carRepairDetailsFxml}`
              );
            if (carRepairDetailsResults.entities.length > 0) {
              const carRepairDetails = carRepairDetailsResults.entities[0];
              var typeOfContact = carRepairDetails["dcar_typeofcontact"];

              if (typeOfContact == 5) {
                saveCurrentPage(formContext);
                alert(
                  "Save and close case, but not update SDM because type of contact is open"
                );
                await saveCurrentPage(formContext, true);
              } else if (typeOfContact == null) {
                saveCurrentPage(formContext);
                alert(
                  "Save case but not update SDM, no value in type of contact"
                );
              } else {
                addToFlexFields(
                  "type_of_contact",
                  carRepairDetails[
                    "dcar_typeofcontact@OData.Community.Display.V1.FormattedValue"
                  ]
                );
                addToFlexFields(
                  "collection_customer",
                  carRepairDetails["dcar_collectioncustomer"]
                );
                addToFlexFields(
                  "collection_address_line_1",
                  carRepairDetails["dcar_collectionaddressline1"]
                );
                addToFlexFields(
                  "collection_address_line_2",
                  carRepairDetails["dcar_collectionaddressline2"]
                );
                addToFlexFields(
                  "collection_address_line_3",
                  carRepairDetails["dcar_collectionaddressline3"]
                );
                addToFlexFields(
                  "collection_postal_code",
                  carRepairDetails["dcar_collectionpostalcode"]
                );
                addToFlexFields(
                  "collection_city",
                  carRepairDetails["dcar_collectioncity"]
                );
                addToFlexFields(
                  "collection_country",
                  carRepairDetails[
                    "dcar_collectioncountry@OData.Community.Display.V1.FormattedValue"
                  ]
                );
                addToFlexFields(
                  "delivery_customer",
                  carRepairDetails["dcar_deliverycustomer"]
                );
                addToFlexFields(
                  "delivery_address_line_1",
                  carRepairDetails["dcar_deliveryaddressline1"]
                );
                addToFlexFields(
                  "delivery_address_line_2",
                  carRepairDetails["dcar_deliveryaddressline2"]
                );
                addToFlexFields(
                  "delivery_address_line_3",
                  carRepairDetails["dcar_deliveryaddressline3"]
                );
                addToFlexFields(
                  "delivery_postal_code",
                  carRepairDetails["dcar_deliverypostalcode"]
                );
                addToFlexFields(
                  "delivery_city",
                  carRepairDetails["dcar_deliverycity"]
                );
                addToFlexFields(
                  "delivery_country",
                  carRepairDetails[
                    "dcar_deliverycountry@OData.Community.Display.V1.FormattedValue"
                  ]
                );
                addToFlexFields(
                  "Inbound Service",
                  incidentRecord[
                    "dcar_inboundservice@OData.Community.Display.V1.FormattedValue"
                  ]
                );
                addToFlexFields(
                  "NFF spec. acces.",
                  carRepairDetails["dcar_depotownership"] == 0
                    ? "REGULAR_PROCESS"
                    : carRepairDetails["dcar_depotownership"] == 1
                    ? "PILOT"
                    : "UNDEFINED"
                );
                addToFlexFields(
                  "system_password",
                  carRepairDetails["dcar_systempassword"]
                );
                var qrStatus = carRepairDetails["dcar_qrstatus"];
                if (qrStatus !== null && qrStatus !== undefined) {
                  addToFlexFields(
                    "QR Status",
                    carRepairDetails[
                      "dcar_qrstatus@OData.Community.Display.V1.FormattedValue"
                    ]
                  );
                } else {
                  alert("QR Status not updated, please Update and save again");
                  return;
                }

                const collectionNotes = carRepairDetails["dcar_collectionnotes"];
                if(collectionNotes !== null || collectionNotes !== undefined) {
                  addNoteToIncident("Collection Notes", collectionNotes, incidentId);
                  addToFlexFields("collection_notes",carRepairDetails["dcar_collectionnotes"]);
                }
               

                var dcar_easybutton = carRepairDetails["dcar_easybutton"];
                if (
                  dcar_easybutton !== null &&
                  (dcar_easybutton === false || dcar_easybutton === true)
                ) {
                  var NFFSpecAccess =
                    dcar_easybutton === false ? "None" : "Easy Button";
                  addToFlexFields("NFF def. freq.", NFFSpecAccess);
                } else {
                  alert(
                    "Incorrect value in dcar_easybutton: " + dcar_easybutton
                  );
                }

                var dcar_lch = carRepairDetails["dcar_lch"];
                switch (dcar_lch) {
                  case 1:
                    var NFFSpecAccess = "LCH ARROW";
                    addToFlexFields("NFF def. freq.", NFFSpecAccess);
                    break;
                  case 2:
                    var NFFSpecAccess = "LCH ARROW ECS";
                    addToFlexFields("NFF def. freq.", NFFSpecAccess);
                    break;
                  case 3:
                    var NFFSpecAccess = "LCH MOOG";
                    addToFlexFields("NFF def. freq.", NFFSpecAccess);
                    break;
                  case 4:
                    var NFFSpecAccess = "LCH GOLDMAN SACHS";
                    addToFlexFields("NFF def. freq.", NFFSpecAccess);
                    break;
                }

                var dcar_kyhd = carRepairDetails["dcar_kyhd"];
                if (
                  dcar_kyhd !== null &&
                  (dcar_kyhd === false || dcar_kyhd === true)
                ) {
                  var KYHD = dcar_kyhd === true ? "YES" : "NO";
                  addToFlexFields("KHYD", KYHD);
                } else {
                  alert("Incorrect value in dcar_kyhd: " + dcar_kyhd);
                }

                var dcar_glsnocourier = carRepairDetails["dcar_glsnocourier"];
                if (
                  dcar_glsnocourier !== null &&
                  (dcar_glsnocourier === false || dcar_glsnocourier === true)
                ) {
                  var glsnocourier = dcar_glsnocourier === true ? "None" : "EMPY_BOX_PIR";
                  addToFlexFields("PS other inf.", glsnocourier);
                } else {
                  alert("Incorrect value in dcar_glsnocourier: " + dcar_glsnocourier);
                }

                const dps_creation = carRepairDetails["dcar_dpscreation"];
                if (dps_creation === true) {
                  addToFlexFields("NFF def. freq.", "DPS Creation");
                }

                var retailer = incidentRecord["_dcar_retailer_value@OData.Community.Display.V1.FormattedValue"];
                if (retailer !== null && retailer !== undefined) 
                {
                  addToFlexFields("Retailer", retailer);
                }

                
                //========================= contact details  ============================//
                const contactId = incidentRecord["_customerid_value"];
                const contactFetchXml = `
                  <fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                  <entity name='contact'>
                  <all-attributes />
                  <filter type='and'>
                  <condition attribute='contactid' operator='eq' value='${contactId}' />
                  </filter>
                  </entity>
                  </fetch>`;
                try {
                  const contactResult =
                    await Xrm.WebApi.retrieveMultipleRecords(
                      "contact",
                      `?fetchXml=${contactFetchXml}`
                    );

                  if (contactResult.entities.length > 0) {
                    const contactRecord = contactResult.entities[0];

                    addToFlexFields(
                      "primary contact",
                      contactRecord["fullname"]
                    );
                    addToFlexFields(
                      "primary contact email",
                      contactRecord["emailaddress1"]
                    );
                    addToFlexFields(
                      "alternative_contact_email",
                      contactRecord["emailaddress2"]
                    );
                    addToFlexFields(
                      "primary contact phone",
                      contactRecord["telephone1"]
                    );
                    addToFlexFields(
                      "alternative_contact_phone",
                      contactRecord["telephone2"]
                    );
                    addToFlexFields(
                      "language code",
                      contactRecord[
                        "dcar_languagecode@OData.Community.Display.V1.FormattedValue"
                      ]
                    );
                  }
                } catch (error) {
                  console.error(
                    "Wystąpił błąd podczas pobierania danych kontaktu:",
                    error
                  );
                }

                //==========================================================================================================//
                //========================= final statement, set status and event, run SDM flow ============================//
                //==========================================================================================================//

                const orderStatus =
                  formContext.getAttribute("dcar_orderstatus");
                const servicestatus =
                  formContext.getAttribute("dcar_servicestatus");
                const serviceevent =
                  formContext.getAttribute("dcar_serviceevent");

                var cad = carRepairDetails["dcar_cad"];
                if (cad !== null && cad === true) {
                  sdmBody.referenceOrderData.OrderStatus = "Firm";
                  const exc_code =
                    formContext.getAttribute("dcar_exceptioncode");
                  exc_code.setValue(35);
                  exc_code.fireOnChange();
                  addToFlexFields(
                    "Exception Code",
                    "CAD - Customer Arranged Date"
                  );

                  var cad_date = carRepairDetails["dcar_customerarrangedate"];
                  var parsedCadDate = await parseCadDate(cad_date);

                  addToFlexFields("revised_collection_date", parsedCadDate);
                  //Firm
                  orderStatus.setValue(1);
                  orderStatus.fireOnChange();
                  sdmBody.referenceOrderData.OrderStatus = "Firm";
                  // Uddate
                  servicestatus.setValue(4);
                  servicestatus.fireOnChange();
                  addToFlexFields(
                    "Service Event",
                    "Customer Arranged Date (CAD)"
                  );
                  //Customer Arranged Date
                  serviceevent.setValue(13);
                  serviceevent.fireOnChange();
                  addToFlexFields("Service Status", "Update");
                } else {
                  //Released
                  orderStatus.setValue(2);
                  orderStatus.fireOnChange();
                  sdmBody.referenceOrderData.OrderStatus = "Released";
                  //Update
                  servicestatus.setValue(4);
                  servicestatus.fireOnChange();
                  addToFlexFields("Service Status", "Update");
                  //Repair Arranged
                  serviceevent.setValue(12);
                  serviceevent.fireOnChange();
                  addToFlexFields("Service Event", "Repair Arranged");
                }
                

                Xrm.Utility.showProgressIndicator(
                  "Update Sdm In progres... please wait"
                );

                callSdm(sdmBody)
                  .then((resp) => {
                    handleSdmResponse(resp);
                  })
                  .catch((error) => {
                    sdmSuccess = false;
                    alert(
                      "UpdateReferenceOrder power automate flow returned an error: " +
                        error.message
                    );
                  })
                  .finally(() => {
                    Xrm.Utility.closeProgressIndicator();
                    if (sdmSuccess) {
                      saveCurrentPage(formContext, true);
                    } else {
                      saveCurrentPage(formContext, false);
                    }
                  });
              }
            }
          } catch (error) {
            alert(
              "Retrive operation for dcar_carrepairdetails with failed result: " +
                error.message
            );
          }
          break;
        // Warranty issue
        case 2:
          var carRepairDetailsid =
            incidentRecord["_dcar_carrepairdetails_value"];
          var carRepairDetailsFxml =
            "<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>" +
            "<entity name='dcar_carrepairdetails'>" +
            "<all-attributes />" +
            "<filter type='and'>" +
            "<condition attribute='dcar_carrepairdetailsid' operator='eq' value='" +
            carRepairDetailsid +
            "' />" +
            "</filter>" +
            "</entity>" +
            "</fetch>";

          try {
            const carRepairDetailsResults =
              await Xrm.WebApi.retrieveMultipleRecords(
                "dcar_carrepairdetails",
                `?fetchXml=${carRepairDetailsFxml}`
              );
            if (carRepairDetailsResults.entities.length > 0) {
              const carRepairDetails = carRepairDetailsResults.entities[0];
              var wiStatus = carRepairDetails["dcar_warrantyissuestatus"];
              const servicestatus =
                formContext.getAttribute("dcar_servicestatus");
              const serviceevent =
                formContext.getAttribute("dcar_serviceevent");

              //wi status 1 is Success
              if (wiStatus == 1) {
                var wiDispodition = carRepairDetails["dcar_widisposition"];
                if (wiDispodition !== null) {
                  addToFlexFields(
                    "WI disposition",
                    carRepairDetails[
                      "dcar_widisposition@OData.Community.Display.V1.FormattedValue"
                    ]
                  );
                  if (wiDispodition == 6) {
                    addToFlexFields("WI description",carRepairDetails["dcar_wiexcdetails"]);
                  } else if (wiDispodition == 2) {
                    addToFlexFields(
                      "delivery_customer",
                      carRepairDetails["dcar_deliverycustomer"]
                    );
                    addToFlexFields(
                      "delivery_address_line_1",
                      carRepairDetails["dcar_deliveryaddressline1"]
                    );
                    addToFlexFields(
                      "delivery_address_line_2",
                      carRepairDetails["dcar_deliveryaddressline2"]
                    );
                    addToFlexFields(
                      "delivery_address_line_3",
                      carRepairDetails["dcar_deliveryaddressline3"]
                    );
                    addToFlexFields(
                      "delivery_postal_code",
                      carRepairDetails["dcar_deliverypostalcode"]
                    );
                    addToFlexFields(
                      "delivery_city",
                      carRepairDetails["dcar_deliverycity"]
                    );
                    addToFlexFields(
                      "delivery_country",
                      carRepairDetails[
                        "dcar_deliverycountry@OData.Community.Display.V1.FormattedValue"
                      ]
                    );
                  }
                  addToFlexFields("WI_DYSP",carRepairDetails["dcar_widysp@OData.Community.Display.V1.FormattedValue"]);
                  var fusion_created_date = carRepairDetails["dcar_offercreateddate"];
                  var parsedFusionDate = await parseCadDate(fusion_created_date);
                  addToFlexFields("FUSION_CREATED_DATE",parsedFusionDate);
                  addToFlexFields("WI_FUSION",carRepairDetails["dcar_offernumber"]);
                  addToFlexFields("QR Status",carRepairDetails["dcar_qrstatus@OData.Community.Display.V1.FormattedValue"]);
                  addToFlexFields("QR Required?",carRepairDetails["dcar_qrrequired@OData.Community.Display.V1.FormattedValue"]);
                  addToFlexFields("NFF spec. acces.",carRepairDetails["dcar_depotownership"] == 0
                      ? "REGULAR_PROCESS"
                      : carRepairDetails["dcar_depotownership"] == 1
                      ? "PILOT"
                      : "UNDEFINED"
                  );
                  //Update
                  addToFlexFields("Service Status", "Update");
                  servicestatus.setValue(4);
                  servicestatus.fireOnChange();
                  //Repair Authorization Granted
                  addToFlexFields(
                    "Service Event",
                    "Repair Authorization Granted"
                  );
                  serviceevent.setValue(17);
                  serviceevent.fireOnChange();
                  const resp = await callSdm(sdmBody);
                  handleSdmResponse(resp);
                  saveCurrentPage(formContext);
                } else {
                  alert(
                    "WI status is Success, but WI disposition is not selected, please make correction and update SDM again"
                  );
                  return;
                }
              } else if(wiStatus === undefined) {
                alert("WI status not set, please choose value and try again.");
                return;
              } else {
                addToFlexFields(carRepairDetails["dcar_widysp@OData.Community.Display.V1.FormattedValue"]);
                var fusion_created_date = carRepairDetails["dcar_offercreateddate"];
                var parsedFusionDate = await parseCadDate(fusion_created_date);
                addToFlexFields("FUSION_CREATED_DATE",parsedFusionDate);
                addToFlexFields("WI_FUSION",carRepairDetails["dcar_offernumber"]);
                addToFlexFields("QR Status",carRepairDetails["dcar_qrstatus@OData.Community.Display.V1.FormattedValue"]);
                addToFlexFields("QR Required?",carRepairDetails["dcar_qrrequired@OData.Community.Display.V1.FormattedValue"]);
                addToFlexFields("NFF spec. acces.",carRepairDetails["dcar_depotownership"] == 0
                    ? "REGULAR_PROCESS"
                    : carRepairDetails["dcar_depotownership"] == 1
                    ? "PILOT"
                    : "UNDEFINED"
                );
                addToFlexFields("system_password",carRepairDetails["dcar_systempassword"]);

                //==========================================================================================================//
                //========================= final statement, set status and event, run SDM flow ============================//
                //==========================================================================================================//

                Xrm.Utility.showProgressIndicator(
                  "Update Sdm In progres... please wait"
                );

                callSdm(sdmBody)
                  .then((resp) => {
                    handleSdmResponse(resp);
                  })
                  .catch((error) => {
                    sdmSuccess = false;
                    alert(
                      "UpdateReferenceOrder power automate flow returned an error: " +
                        error.message
                    );
                  })
                  .finally(() => {
                    Xrm.Utility.closeProgressIndicator();
                    if (sdmSuccess) {
                      saveCurrentPage(formContext, true);
                    } else {
                      saveCurrentPage(formContext, false);
                    }
                  });
              }
            }
          } catch (error) {
            alert(
              "Retrive operation for dcar_carrepairdetails with failed result: " +
                error.message
            );
          }
          break;
        // Customer Experience
        case 3:
          alert(
            "Value in dcar_sectioncategory: " +
              dcar_sectioncategory +
              "Customer Experience"
          );
          break;
        // Customer hold
        case 4:
          var carRepairDetailsid =
            incidentRecord["_dcar_carrepairdetails_value"];
          var carRepairDetailsFxml =
            "<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>" +
            "<entity name='dcar_carrepairdetails'>" +
            "<all-attributes />" +
            "<filter type='and'>" +
            "<condition attribute='dcar_carrepairdetailsid' operator='eq' value='" +
            carRepairDetailsid +
            "' />" +
            "</filter>" +
            "</entity>" +
            "</fetch>";

          try {
            const carRepairDetailsResults =
              await Xrm.WebApi.retrieveMultipleRecords(
                "dcar_carrepairdetails",
                `?fetchXml=${carRepairDetailsFxml}`
              );
            if (carRepairDetailsResults.entities.length > 0) {
              const carRepairDetails = carRepairDetailsResults.entities[0];
              var customerholdstatus =
                carRepairDetails["dcar_customerholdstatus"];
              // Success
              if (customerholdstatus == 1) {
                var additional_information =
                  carRepairDetails["dcar_additionalinformation"];
                if (additional_information != null) {
                  //reset service_question for new value -> set 'none' Hold Optimization
                  addToFlexFields("service_question", "none");
                  addToFlexFields(
                    "additional_information",
                    additional_information
                  );
                  addToFlexFields(
                    "system_password",
                    carRepairDetails["dcar_systempassword"]
                  );
                  addToFlexFields(
                    "QR Status",
                    carRepairDetails[
                      "dcar_qrstatus@OData.Community.Display.V1.FormattedValue"
                    ]
                  );
                  addToFlexFields(
                    "QR Required?",
                    carRepairDetails[
                      "dcar_qrrequired@OData.Community.Display.V1.FormattedValue"
                    ]
                  );
                  addToFlexFields(
                    "NFF spec. acces.",
                    carRepairDetails["dcar_depotownership"] == 0
                      ? "REGULAR_PROCESS"
                      : carRepairDetails["dcar_depotownership"] == 1
                      ? "PILOT"
                      : "UNDEFINED"
                  );
                  //==========================================================================================================//
                  //========================= final statement, set status and event, run SDM flow ============================//
                  //==========================================================================================================//
                  //Update
                  const servicestatus =
                    formContext.getAttribute("dcar_servicestatus");
                  const serviceevent =
                    formContext.getAttribute("dcar_serviceevent");

                  servicestatus.setValue(4);
                  servicestatus.fireOnChange();
                  addToFlexFields("Service Status", "Update");

                  // Check if the attribute exists and has a value
                  if (serviceevent) {
                    const serviceEventValue = serviceevent.getValue();

                    // Check if the value is not null or undefined
                    if (
                      serviceEventValue !== null &&
                      serviceEventValue !== undefined
                    ) {
                      switch (serviceEventValue) {
                        case 18:
                          serviceevent.setValue(14);
                          serviceevent.fireOnChange();
                          addToFlexFields(
                            "Service Event",
                            "Details Provided NFF"
                          );
                          break;
                        case 19:
                          serviceevent.setValue(15);
                          serviceevent.fireOnChange();
                          addToFlexFields(
                            "Service Event",
                            "Details Provided NFF2"
                          );
                          break;
                        case 20:
                          serviceevent.setValue(16);
                          serviceevent.fireOnChange();
                          addToFlexFields(
                            "Service Event",
                            "Details Provided ADM"
                          );
                          break;
                      }
                    }
                  }
                  addNoteToIncident(
                    "Customer Hold Response",
                    additional_information,
                    incidentId
                  );

                  Xrm.Utility.showProgressIndicator(
                    "Update Sdm In progres... please wait"
                  );

                  await callSdm(sdmBody)
                    .then((resp) => {
                      handleSdmResponse(resp);
                    })
                    .catch((error) => {
                      sdmSuccess = false;
                      alert(
                        "UpdateReferenceOrder power automate flow returned an error: " +
                          error.message
                      );
                    })
                    .finally(() => {
                      Xrm.Utility.closeProgressIndicator();
                      if (sdmSuccess) {
                        saveCurrentPage(formContext, true);
                      } else {
                        saveCurrentPage(formContext, false);
                      }
                    });
                } else {
                  alert(
                    "Customer Hold Status is Success, but Additional Information is empty, please make correction and update SDM again"
                  );
                  return;
                }
              }
            }
          } catch (error) {
            alert(
              "Retrive operation for dcar_carrepairdetails with failed result: " +
                error.message
            );
          }
          break;
        // Parts Not Available
        case 5:
          var carRepairDetailsid =
            incidentRecord["_dcar_carrepairdetails_value"];
          var carRepairDetailsFxml =
            "<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>" +
            "<entity name='dcar_carrepairdetails'>" +
            "<all-attributes />" +
            "<filter type='and'>" +
            "<condition attribute='dcar_carrepairdetailsid' operator='eq' value='" +
            carRepairDetailsid +
            "' />" +
            "</filter>" +
            "</entity>" +
            "</fetch>";

          try {
            const carRepairDetailsResults =
              await Xrm.WebApi.retrieveMultipleRecords(
                "dcar_carrepairdetails",
                `?fetchXml=${carRepairDetailsFxml}`
              );
            if (carRepairDetailsResults.entities.length > 0) {
              const carRepairDetails = carRepairDetailsResults.entities[0];

              const dcar_pnacomment = carRepairDetails["dcar_pnacomment"];
              const dcar_status = carRepairDetails["dcar_pnastatus"];
              
              if((dcar_pnacomment != null || dcar_pnacomment != undefined) 
              && (dcar_status != null || dcar_status != undefined)) {

                const pna_status = carRepairDetails["dcar_pnastatus@OData.Community.Display.V1.FormattedValue"]
                addNoteToIncident(
                'PNA Status : ' + pna_status,
                dcar_pnacomment,
                incidentId
              );
              }
            }
          } catch (error) {
            alert(
              "Value in dcar_sectioncategory: " +
                dcar_sectioncategory +
                "Parts Not Available"
            );
          }
          break;
          
          break;
        // Failed Collection
        case 6:
          var carRepairDetailsid =
            incidentRecord["_dcar_carrepairdetails_value"];
          var carRepairDetailsFxml =
            "<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>" +
            "<entity name='dcar_carrepairdetails'>" +
            "<all-attributes />" +
            "<filter type='and'>" +
            "<condition attribute='dcar_carrepairdetailsid' operator='eq' value='" +
            carRepairDetailsid +
            "' />" +
            "</filter>" +
            "</entity>" +
            "</fetch>";

          try {
            const carRepairDetailsResults =
              await Xrm.WebApi.retrieveMultipleRecords(
                "dcar_carrepairdetails",
                `?fetchXml=${carRepairDetailsFxml}`
              );
            if (carRepairDetailsResults.entities.length > 0) {
              const carRepairDetails = carRepairDetailsResults.entities[0];

              const dcar_FailedCollectionResultNote =
                carRepairDetails["dcar_failedcollectionresultnote"];
              const dcar_failedcollectionresult =
                carRepairDetails[
                  "dcar_failedcollectionresult@OData.Community.Display.V1.FormattedValue"
                ];

              addNoteToIncident(
                dcar_failedcollectionresult,
                dcar_FailedCollectionResultNote,
                incidentId
              );
            }
          } catch (error) {
            alert(
              "Retrive operation for dcar_carrepairdetails with failed result: " +
                error.message
            );
          }
          break;
        // Call Back
        case 7:
          var carRepairDetailsid =
            incidentRecord["_dcar_carrepairdetails_value"];
          var carRepairDetailsFxml =
            "<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>" +
            "<entity name='dcar_carrepairdetails'>" +
            "<all-attributes />" +
            "<filter type='and'>" +
            "<condition attribute='dcar_carrepairdetailsid' operator='eq' value='" +
            carRepairDetailsid +
            "' />" +
            "</filter>" +
            "</entity>" +
            "</fetch>";

          try {
            const carRepairDetailsResults =
              await Xrm.WebApi.retrieveMultipleRecords(
                "dcar_carrepairdetails",
                `?fetchXml=${carRepairDetailsFxml}`
              );
            if (carRepairDetailsResults.entities.length > 0) {
              const carRepairDetails = carRepairDetailsResults.entities[0];

              const callbackComment = carRepairDetails["dcar_callbackcomment"];
              const callBackStatusValue =
                carRepairDetails[
                  "dcar_callbackstatus@OData.Community.Display.V1.FormattedValue"
                ];

              addToFlexFields("Call Back Comment", callbackComment);
              addToFlexFields("Call Back Status", callBackStatusValue);
              addNoteToIncident(
                "Call Back Comment",
                callbackComment,
                incidentId
              );

              var dcar_glsnocourier = carRepairDetails["dcar_glsnocourier"];
                if (
                  dcar_glsnocourier !== null &&
                  (dcar_glsnocourier === false || dcar_glsnocourier === true)
                ) {
                  var glsnocourier = dcar_glsnocourier === true ? "None" : "EMPY_BOX_PIR";
                  addToFlexFields("PS other inf.", glsnocourier);
                } else {
                  alert("Incorrect value in dcar_glsnocourier: " + dcar_glsnocourier);
                }

                addToFlexFields(
                  "collection_customer",
                  carRepairDetails["dcar_collectioncustomer"]
                );
                addToFlexFields(
                  "collection_address_line_1",
                  carRepairDetails["dcar_collectionaddressline1"]
                );
                addToFlexFields(
                  "collection_address_line_2",
                  carRepairDetails["dcar_collectionaddressline2"]
                );
                addToFlexFields(
                  "collection_address_line_3",
                  carRepairDetails["dcar_collectionaddressline3"]
                );
                addToFlexFields(
                  "collection_postal_code",
                  carRepairDetails["dcar_collectionpostalcode"]
                );
                addToFlexFields(
                  "collection_city",
                  carRepairDetails["dcar_collectioncity"]
                );
                addToFlexFields(
                  "collection_country",
                  carRepairDetails[
                    "dcar_collectioncountry@OData.Community.Display.V1.FormattedValue"
                  ]
                );
                addToFlexFields(
                  "delivery_customer",
                  carRepairDetails["dcar_deliverycustomer"]
                );
                addToFlexFields(
                  "delivery_address_line_1",
                  carRepairDetails["dcar_deliveryaddressline1"]
                );
                addToFlexFields(
                  "delivery_address_line_2",
                  carRepairDetails["dcar_deliveryaddressline2"]
                );
                addToFlexFields(
                  "delivery_address_line_3",
                  carRepairDetails["dcar_deliveryaddressline3"]
                );
                addToFlexFields(
                  "delivery_postal_code",
                  carRepairDetails["dcar_deliverypostalcode"]
                );
                addToFlexFields(
                  "delivery_city",
                  carRepairDetails["dcar_deliverycity"]
                );
                addToFlexFields(
                  "delivery_country",
                  carRepairDetails[
                    "dcar_deliverycountry@OData.Community.Display.V1.FormattedValue"
                  ]
                );
              
              addToFlexFields("Service Status", "Delayed");
              addToFlexFields("Service Event", "Call Back");
              Xrm.Utility.showProgressIndicator(
                "Update Sdm In progres... please wait"
              );

              callSdm(sdmBody)
                .then((resp) => {
                  handleSdmResponse(resp);
                })
                .catch((error) => {
                  sdmSuccess = false;
                  alert(
                    "UpdateReferenceOrder power automate flow returned an error: " +
                      error.message
                  );
                })
                .finally(() => {
                  Xrm.Utility.closeProgressIndicator();
                  if (sdmSuccess) {
                    saveCurrentPage(formContext, true);
                  } else {
                    saveCurrentPage(formContext, false);
                  }
                });
            }
          } catch (error) {
            alert(
              "Retrive operation for dcar_carrepairdetails with failed result: " +
                error.message
            );
          }

          break;
        //dcar_sectioncategory == default
        case 8:
          var carRepairDetailsid =
            incidentRecord["_dcar_carrepairdetails_value"];
          var carRepairDetailsFxml =
            "<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>" +
            "<entity name='dcar_carrepairdetails'>" +
            "<all-attributes />" +
            "<filter type='and'>" +
            "<condition attribute='dcar_carrepairdetailsid' operator='eq' value='" +
            carRepairDetailsid +
            "' />" +
            "</filter>" +
            "</entity>" +
            "</fetch>";

          try {
            const carRepairDetailsResults =
              await Xrm.WebApi.retrieveMultipleRecords(
                "dcar_carrepairdetails",
                `?fetchXml=${carRepairDetailsFxml}`
              );
            if (carRepairDetailsResults.entities.length > 0) {
              const carRepairDetails = carRepairDetailsResults.entities[0];
              addToFlexFields(
                "collection_customer",
                carRepairDetails["dcar_collectioncustomer"]
              );
              addToFlexFields(
                "collection_address_line_1",
                carRepairDetails["dcar_collectionaddressline1"]
              );
              addToFlexFields(
                "collection_address_line_2",
                carRepairDetails["dcar_collectionaddressline2"]
              );
              addToFlexFields(
                "collection_address_line_3",
                carRepairDetails["dcar_collectionaddressline3"]
              );
              addToFlexFields(
                "collection_postal_code",
                carRepairDetails["dcar_collectionpostalcode"]
              );
              addToFlexFields(
                "collection_city",
                carRepairDetails["dcar_collectioncity"]
              );
              addToFlexFields(
                "collection_country",
                carRepairDetails[
                  "dcar_collectioncountry@OData.Community.Display.V1.FormattedValue"
                ]
              );
              addToFlexFields(
                "delivery_customer",
                carRepairDetails["dcar_deliverycustomer"]
              );
              addToFlexFields(
                "delivery_address_line_1",
                carRepairDetails["dcar_deliveryaddressline1"]
              );
              addToFlexFields(
                "delivery_address_line_2",
                carRepairDetails["dcar_deliveryaddressline2"]
              );
              addToFlexFields(
                "delivery_address_line_3",
                carRepairDetails["dcar_deliveryaddressline3"]
              );
              addToFlexFields(
                "delivery_postal_code",
                carRepairDetails["dcar_deliverypostalcode"]
              );
              addToFlexFields(
                "delivery_city",
                carRepairDetails["dcar_deliverycity"]
              );
              addToFlexFields(
                "delivery_country",
                carRepairDetails[
                  "dcar_deliverycountry@OData.Community.Display.V1.FormattedValue"
                ]
              );
              addToFlexFields(
                "QR Required?",
                carRepairDetails[
                  "dcar_qrrequired@OData.Community.Display.V1.FormattedValue"
                ]
              );
              addToFlexFields(
                "NFF spec. acces.",
                carRepairDetails["dcar_depotownership"] == 0
                  ? "REGULAR_PROCESS"
                  : carRepairDetails["dcar_depotownership"] == 1
                  ? "PILOT"
                  : "UNDEFINED"
              );
              addToFlexFields(
                "system_password",
                carRepairDetails["dcar_systempassword"]
              );
              var qrStatus = carRepairDetails["dcar_qrstatus"];
              if (qrStatus !== null && qrStatus !== undefined) {
                addToFlexFields(
                  "QR Status",
                  carRepairDetails[
                    "dcar_qrstatus@OData.Community.Display.V1.FormattedValue"
                  ]
                );
              } else {
                alert("QR Status not updated, please Update and save again");
                return;
              }

              var dcar_easybutton = carRepairDetails["dcar_easybutton"];
              if (
                dcar_easybutton !== null &&
                (dcar_easybutton === false || dcar_easybutton === true)
              ) {
                var NFFSpecAccess =
                  dcar_easybutton === false ? "None" : "Easy Button";
                addToFlexFields("NFF def. freq.", NFFSpecAccess);
              } else {
                alert("Incorrect value in dcar_easybutton: " + dcar_easybutton);
              }
              //!== undefined
              var dcar_glsnocourier = carRepairDetails["dcar_glsnocourier"];
                if (
                  dcar_glsnocourier !== null &&
                  (dcar_glsnocourier === false || dcar_glsnocourier === true || dcar_glsnocourier === undefined)
                ) {
                  var glsnocourier = dcar_glsnocourier === true ? "None" : "EMPY_BOX_PIR";
                  addToFlexFields("PS other inf.", glsnocourier);
                } else {
                  alert("Incorrect value in dcar_glsnocourier: " + dcar_glsnocourier);
                }
              
              var retailer = incidentRecord["_dcar_retailer_value@OData.Community.Display.V1.FormattedValue"];
                if (retailer !== null && retailer !== undefined) 
                {
                  addToFlexFields("Retailer", retailer);
                }

              //========================= contact details  ============================//
              const contactId = incidentRecord["_customerid_value"];
              const contactFetchXml = `
                <fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>
                <entity name='contact'>
                <all-attributes />
                <filter type='and'>
                <condition attribute='contactid' operator='eq' value='${contactId}' />
                </filter>
                </entity>
                </fetch>`;
              try {
                const contactResult = await Xrm.WebApi.retrieveMultipleRecords(
                  "contact",
                  `?fetchXml=${contactFetchXml}`
                );

                if (contactResult.entities.length > 0) {
                  const contactRecord = contactResult.entities[0];

                  addToFlexFields("primary contact", contactRecord["fullname"]);
                  addToFlexFields(
                    "primary contact email",
                    contactRecord["emailaddress1"]
                  );
                  addToFlexFields(
                    "alternative_contact_email",
                    contactRecord["emailaddress2"]
                  );
                  addToFlexFields(
                    "primary contact phone",
                    contactRecord["telephone1"]
                  );
                  addToFlexFields(
                    "alternative_contact_phone",
                    contactRecord["telephone2"]
                  );
                  addToFlexFields(
                    "language code",
                    contactRecord[
                      "dcar_languagecode@OData.Community.Display.V1.FormattedValue"
                    ]
                  );
                }
              } catch (error) {
                console.error(
                  "Wystąpił błąd podczas pobierania danych kontaktu:",
                  error
                );
              }
              //==========================================================================================================//
              //===================================== final statement, run SDM flow ======================================//
              //==========================================================================================================//

              Xrm.Utility.showProgressIndicator(
                "Update Sdm In progres... please wait"
              );

              await callSdm(sdmBody)
                .then((resp) => {
                  handleSdmResponse(resp);
                })
                .catch((error) => {
                  sdmSuccess = false;
                  alert(
                    "UpdateReferenceOrder power automate flow returned an error: " +
                      error.message
                  );
                })
                .finally(() => {
                  Xrm.Utility.closeProgressIndicator();
                  if (sdmSuccess) {
                    saveCurrentPage(formContext, true);
                  } else {
                    saveCurrentPage(formContext, false);
                  }
                });
            }
          } catch (error) {
            alert(
              "Retrive operation for dcar_carrepairdetails with failed result: " +
                error.message
            );
          }
          break;
        //Cancellation requested
        case 9:
          var serviceEventValue =
            incidentRecord[
              "dcar_serviceevent@OData.Community.Display.V1.FormattedValue"
            ];

          var carRepairDetailsid =
            incidentRecord["_dcar_carrepairdetails_value"];
          var carRepairDetailsFxml =
            "<fetch version='1.0' output-format='xml-platform' mapping='logical' distinct='false'>" +
            "<entity name='dcar_carrepairdetails'>" +
            "<all-attributes />" +
            "<filter type='and'>" +
            "<condition attribute='dcar_carrepairdetailsid' operator='eq' value='" +
            carRepairDetailsid +
            "' />" +
            "</filter>" +
            "</entity>" +
            "</fetch>";

          try {
            const carRepairDetailsResults =
              await Xrm.WebApi.retrieveMultipleRecords(
                "dcar_carrepairdetails",
                `?fetchXml=${carRepairDetailsFxml}`
              );
            if (carRepairDetailsResults.entities.length > 0) {
              const carRepairDetails = carRepairDetailsResults.entities[0];
              const servicestatus =
                formContext.getAttribute("dcar_servicestatus");
              const serviceevent =
                formContext.getAttribute("dcar_serviceevent");
              const serviceEventDb = incidentRecord["dcar_serviceevent"];
              const exc_code = formContext.getAttribute("dcar_exceptioncode");
              servicestatus.setValue(6);
              servicestatus.fireOnChange();
              serviceevent.setValue(serviceEventDb);
              serviceevent.fireOnChange();
              exc_code.setValue(37);
              exc_code.fireOnChange();
              const cancellationnotes =
                carRepairDetails["dcar_cancellationnotes"];
              addToFlexFields("Exception Code", "CNL - Cancelled Call");
              addToFlexFields("Service Status", "Delayed");
              addToFlexFields("Service Event", serviceEventValue);

              addNoteToIncident(
                "Cancelation Notes",
                cancellationnotes,
                incidentId
              );

              Xrm.Utility.showProgressIndicator(
                "Update Sdm In progres... please wait"
              );

              callSdm(sdmBody)
                .then((resp) => {
                  handleSdmResponse(resp);
                })
                .catch((error) => {
                  sdmSuccess = false;
                  alert(
                    "UpdateReferenceOrder power automate flow returned an error: " +
                      error.message
                  );
                })
                .finally(() => {
                  Xrm.Utility.closeProgressIndicator();
                  if (sdmSuccess) {
                    saveCurrentPage(formContext, true);
                  } else {
                    saveCurrentPage(formContext, false);
                  }
                });
            }
          } catch (error) {
            alert(
              "Retrive operation for dcar_carrepairdetails with failed result: " +
                error.message
            );
          }
          break;
        // Incorect Customer Details
        case 10:
          const servicestatus = formContext.getAttribute("dcar_servicestatus");
          const serviceevent = formContext.getAttribute("dcar_serviceevent");
          const exc_code = formContext.getAttribute("dcar_exceptioncode");
          servicestatus.setValue(5);
          servicestatus.fireOnChange();
          serviceevent.setValue(25);
          serviceevent.fireOnChange();
          exc_code.setValue(31);
          exc_code.fireOnChange();

          addToFlexFields("Exception Code", "INC - Incorrect Customer details");
          addToFlexFields("Service Status", "Delayed");
          addToFlexFields("Service Event", "Incorrect Customer details (INC)");

          Xrm.Utility.showProgressIndicator(
            "Update Sdm In progres... please wait"
          );

          callSdm(sdmBody)
            .then((resp) => {
              handleSdmResponse(resp);
            })
            .catch((error) => {
              sdmSuccess = false;
              alert(
                "UpdateReferenceOrder power automate flow returned an error: " +
                  error.message
              );
            })
            .finally(() => {
              Xrm.Utility.closeProgressIndicator();
              if (sdmSuccess) {
                saveCurrentPage(formContext, true);
              } else {
                saveCurrentPage(formContext, false);
              }
            });
          break;
      }
    } else {
      alert("Incident record not found for id: " + incidentId);
    }
  } catch (error) {
    alert(
      "Retrive operation for incident with failed result: " + error.message
    );
  }

  function handleSdmResponse(resp) {
    if (resp.includes("ReferenceOrderResponse")) {
      const referenceOrderIdMatch = resp.match(
        /<ReferenceOrderId>(.*?)<\/ReferenceOrderId>/
      );
      const referenceOrderId = referenceOrderIdMatch
        ? referenceOrderIdMatch[1]
        : null;
      if (referenceOrderId) {
        sdmSuccess = true;
        alert(
          `Order successfully processed. Reference Order ID: ${referenceOrderId} updated in SDM`
        );
      } else {
        sdmSuccess = false;
        alert("Error retrieving Reference Order ID from the response");
      }
    } else if (resp.includes("Error")) {
      const errorMessageMatch = resp.match(/<Error>(.*?)<\/Error>/);
      const errorMessage = errorMessageMatch ? errorMessageMatch[1] : null;
      if (errorMessage) {
        sdmSuccess = false;
        alert(`Error: ${errorMessage}`);
      } else {
        sdmSuccess = false;
        alert(`Error retrieving error message from the response: ${resp}`);
      }
    } else {
      sdmSuccess = false;
      alert("Unknown response received from the SDM server");
    }
  }

  function addToFlexFields(fieldName, fieldValue) {
    if (fieldValue) {
      sdmBody.referenceOrderData.flexFields.push({
        name: fieldName,
        value: fieldValue,
      });
    }
  }
}

async function saveCurrentPage(formContext, closeContext) {
  if (closeContext) {
    try {
      console.log("Change of the form saved");
      const customerField = formContext.getAttribute("_customerid_value");
      if (customerField) {
        const customerValue = customerField.getValue();
        if (customerValue) {
          await customerValue.save();
          console.log("Associated customer (contact) record saved");
        }
      }
      const carDetailsField = formContext.getAttribute(
        "_dcar_carrepairdetails_value"
      );
      if (carDetailsField) {
        const carDetailsValue = carDetailsField.getValue();
        if (carDetailsValue) {
          await carDetailsValue.save();
          console.log("Associated car details record saved");
        }
      }
      // await formContext.data.entity.save("saveandclose");
      await formContext.data.entity.save();
    } catch (error) {
      alert("Failed during save data on this page " + error.message);
    }
  } else {
    try {
      console.log("Change of the form saved");
      const customerField = formContext.getAttribute("_customerid_value");
      if (customerField) {
        const customerValue = customerField.getValue();
        if (customerValue) {
          await customerValue.save();
          console.log("Associated customer (contact) record saved");
        }
      }

      const carDetailsField = formContext.getAttribute(
        "_dcar_carrepairdetails_value"
      );
      if (carDetailsField) {
        const carDetailsValue = carDetailsField.getValue();
        if (carDetailsValue) {
          await carDetailsValue.save();
          console.log("Associated car details record saved");
        }
      }

      await formContext.data.entity.save();
    } catch (error) {
      alert("Failed during save data on this page " + error.message);
    }
  }
}

async function parseCadDate(dateOnlyField) {
  if (dateOnlyField) {
    var date = new Date(dateOnlyField);
    var day = date.getDate().toString().padStart(2, "0");
    var month = (date.getMonth() + 1).toString().padStart(2, "0");
    var year = date.getFullYear();

    var formattedDate = `${year}-${month}-${day}`;
    console.log("Data w formacie yyyy-mm-dd: " + formattedDate);
    return formattedDate;
  }
}


async function parseFusionCreatedDate(dateOnlyField) {
  if (dateOnlyField) {
    var date = new Date(dateOnlyField);
    var day = date.getDate().toString().padStart(2, "0");
    var month = (date.getMonth() + 1).toString().padStart(2, "0");
    var year = date.getFullYear();

    var formattedDate = `${month}-${day}-${year}`;
    console.log("Data w formacie mm-dd-yyyy: " + formattedDate);
    return formattedDate;
  }
}

async function callSdm(sdmBody) {
  try {
    const name = "dcar_sdmupdateribbonvariable";
    const results = await Xrm.WebApi.retrieveMultipleRecords(
      "environmentvariabledefinition",
      `?$filter=schemaname eq '${name}'&$select=environmentvariabledefinitionid&$expand=environmentvariabledefinition_environmentvariablevalue($select=value)`
    );

    if (results.entities.length > 0) {
      const envVariable = results.entities[0];
      jsonData = JSON.parse(
        envVariable.environmentvariabledefinition_environmentvariablevalue[0]
          .value
      );
      const apiUrl = jsonData.apiUrl;
      var myHeaders = new Headers();
      myHeaders.append("Content-Type", "application/json");

      var requestOptions = {
        method: "POST",
        headers: myHeaders,
        body: JSON.stringify(sdmBody),
        redirect: "follow",
      };
      const response = await fetch(apiUrl, requestOptions);
      if (!response.ok) {
        throw new Error("Network response was not ok");
      }

      const result = await response.text();
      return result;
    }
  } catch (error) {
    console.error("Error:", error.message);
    throw error;
  }
}

// async function callSdm(sdmBody) {
//   try {
//     const name = "dcar_sdmupdateribbonvariable";
//     const results = await Xrm.WebApi.retrieveMultipleRecords(
//       "environmentvariabledefinition",
//       `?$filter=schemaname eq '${name}'&$select=environmentvariabledefinitionid&$expand=environmentvariabledefinition_environmentvariablevalue($select=value)`
//     );

//     if (results.entities.length > 0) {
//       const envVariable = results.entities[0];
//       const jsonData = JSON.parse(
//         envVariable.environmentvariabledefinition_environmentvariablevalue[0]
//           .value
//       );

//       const clientId = jsonData.clientId;
//       const clientSecret = jsonData.clientSecret;
//       const tokenUrl = jsonData.tokenUrl;
//       const apiUrl = jsonData.apiUrl;
//       const resource = jsonData.resource;

//       const tokenResponse = await fetch(
//         tokenUrl,
//         requestOptions
//       );

//       if (!tokenResponse.ok) {
//         throw new Error("Failed to obtain access token");
//       }

//       const tokenResult = await tokenResponse.json();
//       const accessToken = tokenResult.access_token;

//       const apiResponse = await fetch(apiUrl, {
//         method: "POST",
//         headers: {
//           "Content-Type": "application/json",
//           Authorization: `Bearer ${accessToken}`,
//         },
//         body: JSON.stringify(sdmBody),
//       });

//       if (!apiResponse.ok) {
//         throw new Error("API request failed");
//       }

//       const apiResult = await apiResponse.text();
//       return apiResult;
//     } else {
//       console.error("No environment variable found with the name:", name);
//     }
//   } catch (error) {
//     console.error("An error occurred:", error.message);
//     throw error;
//   }
// }

function addNoteToIncident(subject, notetext, incidentId) {
  incidentId = incidentId.replace(/[{}]/g, "");
  var data = {
    subject: subject,
    notetext: notetext,
    "objectid_incident@odata.bind": "/incidents(" + incidentId + ")",
  };

  Xrm.WebApi.createRecord("annotation", data).then(
    function success(result) {
      console.log("Note created with ID: " + result.id);
    },
    function (error) {
      console.log("Note not created: " + error.message);
      // Handle error conditions
    }
  );
}
