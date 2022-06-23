const ExcelJS = require('exceljs');
const xlsx = require("xlsx");
const { carriers } = require('../static/carriers');
const { headers } = require('../static/headers');

const createCarrierAvailablityFile = async (carriersList, keys) => {
    let workbook = xlsx.readFile(process.cwd() + "/" + "./files/BoxcheckRegionalCarriers.xlsx", { cellText: false, raw: true, cellDates: true });
    let sheetNames = workbook.SheetNames;
    let rawCarriersData = xlsx.utils.sheet_to_json(workbook.Sheets[sheetNames[0]], { defval: "" });
    let modifiedCarriersData = [];
    for (let rawCarriers of rawCarriersData) {
        if (Object.keys(rawCarriers).length > 0) {
            let getRequiredCarrier = carriersList.find(list => list["additionalCarrierID"] == rawCarriers.Carrier);
            if (typeof getRequiredCarrier !== "undefined" && rawCarriers.Carrier == getRequiredCarrier.additionalCarrierID) {
                modifiedCarriersData.push({
                    "Origin State": rawCarriers.State,
                    "Origin Zip": rawCarriers.Zip,
                    "Destination State": rawCarriers.State,
                    "Destination Zip": rawCarriers.Zip,
                    "Carrier Name": getRequiredCarrier.fullName,
                    "Carrier ID": getRequiredCarrier.carrierId,
                    "Carrier Service": getRequiredCarrier.service,
                    "Product Types": getRequiredCarrier.productTypes,
                    "Rate": getRequiredCarrier.Rate
                })
            }
        }
    }
    let header = keys.map((key, index) => {
        return {
            header: key,
            key: key,
            width: 20
        };
    });

    let workBook = new ExcelJS.Workbook();
    let workSheet = workBook.addWorksheet("Carrier Availablity");
    workSheet.columns = header;
    workSheet.addRows(modifiedCarriersData);
    let filePath = "./files/CarrierAvailablity.xlsx"
    await workBook.xlsx.writeFile(filePath)
}
createCarrierAvailablityFile(carriers, headers);