async function wrapper2(sheetNameNew, sourceSheet, sourceTable, variable, value) {
    await clearFilters(sourceSheet, sourceTable);
    await filterTable(sourceSheet, sourceTable, variable, value);
    await copyVisibleRange(sourceSheet, sourceTable, sheetNameNew);    
}


async function filterTable(worksheet, sourceTable, variable, value) {
    await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getItem(worksheet);
        const table = sheet.tables.getItem(sourceTable);
        let filter = table.columns.getItem(variable).filter;
        filter.apply({
            filterOn: Excel.FilterOn.values,
            values: [value]
        });
        await context.sync();
    });
}

async function clearFilters(worksheet, sourceTable) {
    await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getItem(worksheet);
        const table = sheet.tables.getItem(sourceTable);
        table.clearFilters();
        await context.sync();
    });
}


async function copyVisibleRange2(worksheetSource, tableSource, worksheetDest) {
    await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getItem(worksheetSource);
        const table = sheet.tables.getItem(tableSource);
        const visibleRange = table.getDataBodyRange().getVisibleView().load("values");
        visibleRange.load("address");
        await context.sync();        
        context.workbook.worksheets.getItemOrNullObject(worksheetDest).delete();
        const sheetDest = context.workbook.worksheets.add(worksheetDest);
        sheetDest.getRange("A1").copyFrom(visibleRange);
        await context.sync();
    });
}

async function copyVisibleRange(worksheetSource, tableSource, worksheetDest) {
    await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getItem(worksheetSource);
        const table = sheet.tables.getItem(tableSource);
        const visibleRange = table.getRange().getVisibleView().load("values");
        await context.sync();
        
        let values = visibleRange.values;
        let rowCount = values.length;
        let columnCount = values[0].length;

        context.workbook.worksheets.getItemOrNullObject(worksheetDest).delete();
        let sheetDest = context.workbook.worksheets.add(worksheetDest);
        let range = sheetDest.getRangeByIndexes(0, 0, rowCount, columnCount);        
        range.values = values;
        sheetDest.getUsedRange().format.autofitColumns();
        sheetDest.getUsedRange().format.autofitRows();
       
        let newTable = sheetDest.tables.add(range, true);
        newTable.name = worksheetDest;
        await context.sync();
    });
}

//top!
async function copyTableToAnotherSheet(){
    await Excel.run(async (context) => {
        let sheet1 = context.workbook.worksheets.getItem("Tabelle1");
        const table = sheet1.tables.getItem("Tabelle1");
        let range = table.getRange();
        range.load("address");
        await context.sync();

        let sheet2 = context.workbook.worksheets.getItem("Tabelle3");
        sheet2.getRange("A1").copyFrom(range);
        await context.sync();
    });
}

//top!
async function getValuesFromColumn(worksheetSource, tableSource, column) {
    return new Office.Promise(async function (resolve) {
        await Excel.run(async (context) => {
            let sheet = context.workbook.worksheets.getItem(worksheetSource);
            const table = sheet.tables.getItem(tableSource);
            const columnRange = table.columns.getItem(column).getDataBodyRange().load("values");
            await sheet.context.sync();
            const columnValues = columnRange.values;
            await context.sync();
            resolve(columnValues);
        });
    })
   
}


async function copyTableHeaders(tableName, sheetName) {
    await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getItem(sheetName);
        const targetCell = sheet.getRange("A1");
        targetCell.formulas = [["="+tableName+"[#Headers]"]];
        const spillRange = targetCell.getSpillingToRange();
        spillRange.load("address");
        sheet.getUsedRange().format.autofitColumns();
        await context.sync();
    });
}

async function applyFilterFunction(sheetname, table, variable, value) {
    await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getItem(sheetname);
        const targetCell = sheet.getRange("A2");
        targetCell.formulas = [
            ['=FILTER('+table+'[#All],' +table + '[[#All],['+variable+']]="'+value+'", "")']
            //['=FILTER(Tabelle1[#All], Tabelle1[[#All],[A]]="1", "")']
        ];
        const spillRange = targetCell.getSpillingToRange();
        spillRange.load("address");
        sheet.getUsedRange().format.autofitColumns();
        await context.sync();
    });
}

async function createSheet(sheetName) {
    await Excel.run(async (context) => {
        context.workbook.worksheets.getItemOrNullObject(sheetName).delete();
        const sheet = context.workbook.worksheets.add(sheetName);
        sheet.load("name, position");
        await context.sync();
    });
}

async function listWorksheets(dotNetReference) {
    await Office.onReady();
    await Excel.run(async (context) => {
        const sheets = context.workbook.worksheets;
        sheets.load("items/name");
        await context.sync();
        let allSheets = [];

        for (let i in sheets.items) {
            const tables = sheets.items[i].tables;
            tables.load('name, count, headers, columns')
            await context.sync();
            let allTables = []
            for (let j in tables.items) {
                let tableheaders = tables.items[j].columns.items;
                let alltableheaders = []
                for (let k in tableheaders)
                {
                    alltableheaders.push(tableheaders[k].name);
                }
                allTables.push({ tablename: tables.items[j].name, categories: alltableheaders });

            }
            allSheets.push({sheetname: sheets.items[i].name, tables: allTables});
        }
        dotNetReference.invokeMethodAsync("CallbackAllWorksheets", allSheets);
    });
}