const ExcelJS = require('exceljs');
const xlsx = require("xlsx");
const fs = require('fs');
const path = require('path');

const createCarrierAvailablityFile = async (carriersList, keys) => {
    let workbook = xlsx.readFile(process.cwd() + "/" + "./files/BoxcheckRegionalCarriers.xlsx", { cellText: false, raw: true, cellDates: true });
    let sheetNames = workbook.SheetNames;
    let rawCarriersData = xlsx.utils.sheet_to_json(workbook.Sheets[sheetNames[0]], { defval: "" });
    let modifiedCarriersData = [];
    for (let rawCarriers of rawCarriersData) {
        if (Object.keys(rawCarriers).length > 0) {
            let getRequiredCarrier = carriersList.find(list => list["carrierId"] == rawCarriers.Carrier);
            if (typeof getRequiredCarrier !== "undefined" && rawCarriers.Carrier == getRequiredCarrier.carrierId) {
                modifiedCarriersData.push({
                    "Origin State": rawCarriers.State,
                    "Origin Zip": rawCarriers.Zip,
                    "Destination State": rawCarriers.State,
                    "Destination Zip": rawCarriers.Zip,
                    "Carrier Name": getRequiredCarrier.fullName,
                    "Carrier ID": rawCarriers.Carrier,
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

let carriers = [
    {
        service: "ST",
        carrierId: "PRMD",
        fullName: "ProMed Delivery",
        productTypes: "beer,spirit,wine,vape",
        Rate: "20.00"

    },
    {
        service: "ST",
        carrierId: "ZIP",
        fullName: "Zip Express",
        productTypes: "beer,spirit,vape",
        Rate: "20.00"

    },
    {
        service: "ST",
        carrierId: "MERC",
        fullName: "Mercury",
        productTypes: "beer,spirit,wine,vape",
        Rate: "20.00"

    },
    {
        service: "ST",
        carrierId: "STAT",
        fullName: "STAT Overnight",
        productTypes: "beer,spirit,wine,vape",
        Rate: "11.99"

    },
    {
        service: "ST",
        carrierId: "SONIC",
        fullName: "Sonic Systems Inc",
        productTypes: "beer,spirit,wine,vape",
        Rate: "20.00"

    },
    {
        service: "ST",
        carrierId: "GS",
        fullName: "Granite State Shuttle",
        productTypes: "vape",
        Rate: "20.00"

    },
    {
        service: "ST",
        carrierId: "JTL",
        fullName: "JET Transportation & Logistics",
        productTypes: "vape",
        Rate: "20.00"

    },
    {
        service: "ST",
        carrierId: "AEX",
        fullName: "American Eagle Express",
        productTypes: "vape",
        Rate: "20.00"

    },
    {
        service: "ST",
        carrierId: "GLS",
        fullName: "General Logistics Systems",
        productTypes: "beer,spirit,wine",
        Rate: "20.00"

    },
    {
        service: "ST",
        carrierId: "DI",
        fullName: "DeliverIT",
        productTypes: "beer,spirit,wine,vape",
        Rate: "20.00"

    },

]
let keys = ["Origin State", "Origin Zip", "Destination State", "Destination Zip", "Carrier Name", "Carrier ID",
    "Carrier Service", "Product Types", "Rate"]

createCarrierAvailablityFile(carriers, keys);