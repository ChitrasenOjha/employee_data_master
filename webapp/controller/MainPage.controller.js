sap.ui.define([
    "sap/ui/core/mvc/Controller",
    "sap/ui/model/json/JSONModel",
    "employeedatamaster/util/xlsx.full.min",
    "sap/m/MessageBox",
    "sap/m/MessageToast"
],
    function (Controller, JSONModel, xlsx, MessageBox, MessageToast) {
        "use strict";
        var TO_ITEMS = [];
        var uploadedCount = 0;
        var employeeCount = 0;
        var isTemplateValid = true;
        return Controller.extend("employeedatamaster.controller.MainPage", {
            onInit: function ()
            {

                isTemplateValid = true;
                
                //-----------Combox logic----------

                // Promise.all([
                //     this.loadJSON("/model/countries.json"),
                //     this.loadJSON("/model/verticals.json")
                // ]).then(function (aData) {
                //     this.allVerticals = aData[1];
                //     var oModel = new sap.ui.model.json.JSONModel({
                //         countries: aData[0],
                //         filteredVerticals: []
                //     });
                //     this.getView().setModel(oModel, "local");

                // }.bind(this)).catch(function (err) {
                //     console.error("Error loading JSON:", err);
                // });

                //CSRF Token logic

                this._csrfToken = null;
                fetch("/sap/opu/odata/sap/ZR_VALIDATION_SF_SRV/", {
                    method: "GET",
                    headers: { "x-csrf-token": "fetch" },
                    credentials: "include"
                }).then(response => {
                    this._csrfToken = response.headers.get("x-csrf-token");
                    console.log("CSRF Token:", this._csrfToken);

                }).catch(err => console.error("CSRF fetch failed", err));
            },

            //-------------JSON Loader Helper---------------

            // loadJSON: function (sPath)
            // {
            //     return new Promise(function (resolve, reject) {
            //         var oModel = new sap.ui.model.json.JSONModel();
            //         oModel.attachRequestCompleted(function () {
            //             resolve(oModel.getData());
            //         });
            //         oModel.attachRequestFailed(function () {
            //             reject("Failed to load: " + sPath);
            //         });
            //         oModel.loadData(sPath);
            //     });
            // },


            //-------------------On Country slection ----------------------

            // onCountryChange: function (oEvent)
            // {
            //     var sCountryId = oEvent.getSource().getSelectedKey();
            //     var aFiltered = this.allVerticals.filter(function (v) {
            //         return v.countryId === sCountryId;
            //     });
            //     var oModel = this.getView().getModel("local");
            //     oModel.setProperty("/filteredVerticals", aFiltered);
            //     this.byId("cbVertical").setEnabled(true);
            // },
            onTemplateSelection: function (oEvent) {
                var oGroup = oEvent.getSource();
                var iIndex = oEvent.getParameter("selectedIndex");
                if (iIndex === -1) {
                    this.selectedFileTemplate = null;
                    return;
                }
                var sSelectedValue = oGroup.getButtons()[iIndex].getText();
                this.selectedFileTemplate = sSelectedValue;
                this.checkEnableValidateButton();
            },


            onEmployeeCountChange: function (oEvent)
            {
                var oInput = oEvent.getSource();
                var sValue = oInput.getValue();
                if (!sValue)
                {
                    oInput.setValueState("Error");
                    oInput.setValueStateText("Employee count is required");
                    this.employeeCount = null;
                    this.checkEnableValidateButton();
                    return;
                }
                var iCount = parseInt(sValue, 10);
                if (isNaN(iCount) || iCount <= 0)
                {
                    oInput.setValueState("Error");
                    oInput.setValueStateText("Employee count must be greater than 0");
                    this.employeeCount = null;
                    this.checkEnableValidateButton();
                }
                else
                {
                    oInput.setValueState("None");
                    this.employeeCount = iCount;
                    this.checkEnableValidateButton();
                }
            },
            onDateSelection: function (oEvent) {
                var oDateValue = oEvent.getSource().getValue();
                this.selectedDate = oDateValue;
                this.checkEnableValidateButton();
            },
           
            onFileChange: function (oEvent) {
                var oFileUploader = this.byId("fileUploader");
                this._file = oEvent.getParameter("files")[0];
                if (this._file) {
                    oFileUploader.setValue(this._file.name);
                    //this.checkEnableValidateButton();
                }
            },


        //--------------------------ACTUAL VALIDATION lOGIC----------------------------------

            onValidate: function () {
                var oModel = this.getOwnerComponent().getModel();
                var oFileUploader = this.byId("fileUploader");
                var file = oFileUploader.getFocusDomRef().files[0];
                if (!this.selectedFileTemplate) {
                    MessageToast.show("Please select a template first!");
                    return;
                }
                if (!file) {
                    MessageToast.show("Please select a file first!");
                    return;
                }
                var file1 = file;
                var excelData =
                {
                    "RuleFieldID": "MTSL",
                    "TemplateId": this.selectedFileTemplate,
                    "FileName": file1.name?.toString() || "",
                    "NoOfEmps": this.employeeCount?.toString() || "",
                    "CutoffDate": this.selectedDate?.toString() || "",
                    "TO_ITEMS": TO_ITEMS
                }

                oFileUploader.addHeaderParameter(new sap.ui.unified.FileUploaderParameter({
                    name: "slug",
                    value: file.name
                }));

                oFileUploader.addHeaderParameter(new sap.ui.unified.FileUploaderParameter({
                    name: "x-csrf-token",
                    value: this._csrfToken
                }));


                if (!this.employeeCount) {
                    MessageToast.show("Please enter employee count!");
                    return;
                }
                if (uploadedCount === 0) {
                    MessageToast.show("Please upload file first");
                    return;
                }
                if (uploadedCount !== this.employeeCount)
                {
                    MessageBox.warning(
                        "Data Mismatch!\n\nEntered Count: " +
                        this.employeeCount +
                        "\nUploaded Records: " +
                        uploadedCount
                    );
                    this.resetValues();
                    return;
                }

                // Fire upload
                oFileUploader.upload();
                sap.ui.core.BusyIndicator.show(0);
                var that = this;
                oModel.create("/zemp_headerSet", excelData, {
                    success: function (oResponse) {

                        sap.ui.core.BusyIndicator.hide();
                        MessageToast.show("File uploaded successfully!");


                        // that.getView().getModel("validatedData").setData(oResponse);
                        // var data = that.getView().getModel("validatedData").getData();
                        // console.log("Validated Data:", data);

                        that.onDownloadValidatedExcel(oResponse.TO_ITEMS.results);
                        var oFU = that.byId("fileUploader");
                        oFU.clear();
                        if (oFU._oFileUpload) {
                            oFU._oFileUpload.value = "";
                        }
                        isTemplateValid=false;
                        that.checkEnableValidateButton();
                    },
                    error: function (oError) {
                        sap.ui.core.BusyIndicator.hide();
                        MessageBox.error("Upload failed: " + oError.message);
                        var oFU = that.byId("fileUploader");
                        oFU.clear();
                        if (oFU._oFileUpload) {
                            oFU._oFileUpload.value = "";
                        }
                        isTemplateValid=false;
                        that.checkEnableValidateButton();
                    }
                });
            },
            onFileUpload: function () {
                var oFileUploader = this.byId("fileUploader");
                //sap.ui.core.BusyIndicator.show(0);

                var file = oFileUploader.getFocusDomRef().files[0];
                
                TO_ITEMS = [];
                uploadedCount = 0;

                var oModel = this.getOwnerComponent().getModel();
                var that = this;

                var reader = new FileReader();
                reader.onload = function (e) {

                    var data = new Uint8Array(e.target.result);
                    var workbook = XLSX.read(data, { type: "array" });
                    var worksheet = workbook.Sheets[workbook.SheetNames[0]];

                    var range = XLSX.utils.decode_range(worksheet["!ref"]);
                    range.s.r = 5;
                    range.s.c = 0;

                    const rowsAsArrays = XLSX.utils.sheet_to_json(worksheet, {
                        header: 1,
                        defval: "",
                        range: range,
                        blankrows: true
                    });

                    const headerRow = rowsAsArrays[9];
                    const columnCount = headerRow.filter(cell => cell !== "").length;
                    console.log(headerRow);
                    console.log("No of Columns:", columnCount);

                    if (columnCount != 109) {
                        MessageToast.show("Incorrect File Template ! Upload Corrrect File Template");
                        sap.ui.core.BusyIndicator.hide();
                        isTemplateValid = false;
                        that.checkEnableValidateButton();
                        return;
                    }
                    isTemplateValid=true;
                    that.checkEnableValidateButton();
                    uploadedCount = rowsAsArrays.length - 10;
                    rowsAsArrays.forEach(row => {

                        var data = {
                            RuleFieldID: row[0]?.toString() || "",
                            Ha001: row[1]?.toString() || "",
                            Bu003: row[2]?.toString() || "",
                            Bi002: row[3]?.toString() || "",
                            Bi003: row[4]?.toString() || "",
                            Bi004: row[5]?.toString() || "",
                            Ei001: row[6]?.toString() || "",
                            Ei002: row[7]?.toString() || "",
                            Ei007: row[8]?.toString() || "",
                            Ei004: row[9]?.toString() || "",
                            Ei003: row[10]?.toString() || "",
                            Jh001: row[11]?.toString() || "",
                            Jm001: row[12]?.toString() || "",
                            Pi017: row[13]?.toString() || "",
                            Pi001: row[14]?.toString() || "",
                            Pi003: row[15]?.toString() || "",
                            Pi014: row[16]?.toString() || "",
                            Pi007: row[17]?.toString() || "",
                            Pi006: row[18]?.toString() || "",
                            Pi009: row[19]?.toString() || "",
                            Pi010: row[20]?.toString() || "",
                            Pi011: row[21]?.toString() || "",
                            Pi002: row[22]?.toString() || "",
                            Pi004: row[23]?.toString() || "",
                            Pi013: row[24]?.toString() || "",
                            Pi005: row[25]?.toString() || "",
                            Pi019: row[26]?.toString() || "",
                            Pi015: row[27]?.toString() || "",
                            Pi008: row[28]?.toString() || "",
                            Bm001: row[29]?.toString() || "",
                            Bm004: row[30]?.toString() || "",
                            Bm002: row[31]?.toString() || "",
                            Bm005: row[32]?.toString() || "",
                            Bm003: row[33]?.toString() || "",
                            Pm001: row[34]?.toString() || "",
                            Pm004: row[35]?.toString() || "",
                            Pm002: row[36]?.toString() || "",
                            Pm005: row[37]?.toString() || "",
                            Pm003: row[38]?.toString() || "",
                            Eb001: row[39]?.toString() || "",
                            Eb002: row[40]?.toString() || "",
                            Ep001: row[41]?.toString() || "",
                            Ep002: row[42]?.toString() || "",
                            Na001: row[43]?.toString() || "",
                            Na002: row[44]?.toString() || "",
                            Na003: row[45]?.toString() || "",
                            Na004: row[46]?.toString() || "",
                            Nb001: row[47]?.toString() || "",
                            Nb002: row[48]?.toString() || "",
                            Nb003: row[49]?.toString() || "",
                            Nb004: row[50]?.toString() || "",
                            Nc001: row[51]?.toString() || "",
                            Nc002: row[52]?.toString() || "",
                            Nc003: row[53]?.toString() || "",
                            Nc004: row[54]?.toString() || "",
                            Ec001: row[55]?.toString() || "",
                            Ec002: row[56]?.toString() || "",
                            Ec003: row[57]?.toString() || "",
                            Ec004: row[58]?.toString() || "",
                            Ec006: row[59]?.toString() || "",
                            Ec008: row[60]?.toString() || "",
                            Ec007: row[61]?.toString() || "",
                            Da012: row[62]?.toString() || "",
                            Da001: row[63]?.toString() || "",
                            Da006: row[64]?.toString() || "",
                            Da002: row[65]?.toString() || "",
                            Da003: row[66]?.toString() || "",
                            Da008: row[67]?.toString() || "",
                            Da009: row[68]?.toString() || "",
                            Da007: row[69]?.toString() || "",
                            Da011: row[70]?.toString() || "",
                            Db012: row[71]?.toString() || "",
                            Db001: row[72]?.toString() || "",
                            Db006: row[73]?.toString() || "",
                            Db002: row[74]?.toString() || "",
                            Db003: row[75]?.toString() || "",
                            Db008: row[76]?.toString() || "",
                            Db009: row[77]?.toString() || "",
                            Db007: row[78]?.toString() || "",
                            Db011: row[79]?.toString() || "",
                            Jc032: row[80]?.toString() || "",
                            Jc031: row[81]?.toString() || "",
                            Jc005: row[82]?.toString() || "",
                            Jc015: row[83]?.toString() || "",
                            Jc025: row[84]?.toString() || "",
                            Jc043: row[85]?.toString() || "",
                            Jc021: row[86]?.toString() || "",
                            Jc011: row[87]?.toString() || "",
                            Jc035: row[88]?.toString() || "",
                            Jc054: row[89]?.toString() || "",
                            Jc061: row[90]?.toString() || "",
                            Jc006: row[91]?.toString() || "",
                            Jc012: row[92]?.toString() || "",
                            Jc036: row[93]?.toString() || "",
                            Jc045: row[94]?.toString() || "",
                            Jc003: row[95]?.toString() || "",
                            Jc008: row[96]?.toString() || "",
                            Jc004: row[97]?.toString() || "",
                            Jc030: row[98]?.toString() || "",
                            Jc013: row[99]?.toString() || "",
                            Jc029: row[100]?.toString() || "",
                            Jc007: row[101]?.toString() || "",
                            Jc047: row[102]?.toString() || "",
                            Jc027: row[103]?.toString() || "",
                            Jc019: row[104]?.toString() || "",
                            Jc049: row[105]?.toString() || "",
                            Jc051: row[106]?.toString() || "",
                            Jc056: row[107]?.toString() || "",
                            Jc044: row[108]?.toString() || ""
                        };

                        TO_ITEMS.push(data);
                    });
                    //sap.ui.core.BusyIndicator.hide();
                    //MessageToast.show("File uploaded successfully!");
                }
                reader.readAsArrayBuffer(file);
                this.checkEnableValidateButton();
            },
            onDownloadValidatedExcel: function (oResponse) {

                if (!oResponse || !oResponse.length) {
                    MessageToast.show("Invalid file format. Please upload the correct Excel template.");
                    return;
                }

                // var hasError = oResponse.some(function (item) {
                //     return item.REMARKS && item.REMARKS.toLowerCase() !== "ok";
                // });

                // if (!hasError) {
                //    MessageToast.show("File is correct. All records validated successfully.");
                //     return;
                // }

                var reorderedList = oResponse.map(function (item) {
                    var copy = Object.assign({}, item);
                    delete copy.__metadata;

                    var reordered = {};
                    if (copy.REMARKS !== undefined) {
                        reordered.REMARKS = copy.REMARKS;
                    }
                    Object.keys(copy).forEach(function (key) {
                        if (key !== "REMARKS") {
                            reordered[key] = copy[key];
                        }
                    });

                    return reordered;
                });

                var worksheet = XLSX.utils.json_to_sheet(reorderedList);
                var workbook = XLSX.utils.book_new();
                XLSX.utils.book_append_sheet(workbook, worksheet, "Validated Data");

                var excelBinary = XLSX.write(workbook, {
                    bookType: "xlsx",
                    type: "array"
                });

                var blob = new Blob([excelBinary], {
                    type: "application/octet-stream"
                });

                var url = URL.createObjectURL(blob);
                var a = document.createElement("a");
                a.href = url;
                a.download = "Validated_Sheet.xlsx";
                a.click();
                URL.revokeObjectURL(url);

                MessageToast.show("Excel downloaded successfully!");
            },

            checkEnableValidateButton: function ()
            {
                var oFileUploader = this.byId("fileUploader");
                var file = oFileUploader.getFocusDomRef().files[0];
                
                var validateEnable = !!this.selectedFileTemplate &&
                    !! this.selectedDate &&
                    !! file &&
                    !!this.employeeCount &&
                    isTemplateValid === true; 
                this.byId("_IDGenButton1").setEnabled(validateEnable);

            },
            
            onClearPress: function () {

                var oFU = this.byId("fileUploader");
                var file = oFU.getFocusDomRef().files[0];
                if(file===undefined)
                {
                    MessageToast.show("No File selected!");
                    return;
                }
                oFU.clear();
                if (oFU._oFileUpload) {
                    oFU._oFileUpload.value = "";
                }
                MessageToast.show("File selection cleared successfully!");
                this.isTemplateValid = false;
                this.checkEnableValidateButton();

            }
        });
    });
