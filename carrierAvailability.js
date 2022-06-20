const ExcelJS = require('exceljs');
const xlsx = require("xlsx");
const fs = require('fs');
const path = require('path');

const createCarrierAvailablityFile = async (carriersList, keys) => {
    let workbook = xlsx.readFile(process.cwd() + "/" + "./files/BoxcheckRegionalCarriers.xlsx", { cellText: false, raw: true, cellDates: true });
    let sheetNames = workbook.SheetNames;
    let rawCarriersData = xlsx.utils.sheet_to_json(workbook.Sheets[sheetNames[0]], { defval: "" });
    let modifiedCarriersData = [];
    let count = 0;
    for (let rawCarriers of rawCarriersData) {
        if (Object.keys(rawCarriers).length > 0) {
            let getRequiredCarrier = carriersList.find(list => list["name"] == rawCarriers.Carrier);
            count++
            if (typeof getRequiredCarrier !== "undefined" && rawCarriers.Carrier == getRequiredCarrier.name) {
                modifiedCarriersData.push({
                    "Origin State": rawCarriers.State,
                    "Origin Zip": rawCarriers.Zip,
                    "Destination State": rawCarriers.State,
                    "Destination Zip": rawCarriers.Zip,
                    "Carrier Name": rawCarriers.Carrier,
                    "Carrier Service": getRequiredCarrier.service,
                    "Product Types": getRequiredCarrier.productTypes
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

let carriers = [
    {
        name: "PROMED",
        service: "ST",
        fullName: "ProMed",
        productTypes: ["beer", "spirit", "vape", "wine"]
    },
    {
        name: "ZIP",
        service: "ST",
        fullName: "ZIP",
        productTypes: ["beer", "spirit", "vape"]
    },
    {
        name: "MERC",
        service: "ST",
        fullName: "Mercury",
        productTypes: ["beer", "spirit", "vape", "wine"]
    },
    {
        name: "STAT",
        service: "ST",
        fullName: "STAT Overnight",
        productTypes: ["beer", "spirit", "vape", "wine"]
    },
    {
        name: "SONIC",
        service: "ST",
        fullName: "SONIC",
        productTypes: ["beer", "spirit", "vape", "wine"]
    },
    {
        name: "Granite State Shuttle",
        service: "ST",
        fullName: "Granite State Shuttle",
        productTypes: ["vape"]
    },
    {
        name: "JET",
        service: "ST",
        fullName: "JET Transportation & Logistics",
        productTypes: ["vape"]
    },
    {
        name: "Deliver-It",
        service: "ST",
        fullName: "Deliver IT",
        productTypes: ["beer", "spirit", "wine"]
    },

]
let keys = ["Origin State", "Origin Zip", "Destination State", "Destination Zip", "Carrier Name", "Carrier Service", "Product Types"]

createCarrierAvailablityFile(carriers, keys);