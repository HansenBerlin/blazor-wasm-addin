async function wrapper2(sourceSheet, sourceTable, variable, v) {
    await Office.onReady();
    return new Office.Promise(async function (resolve) {
        await Excel.run(async (context) => {
            for (let j = 0; j < v.length; j++) {
                console.log(v[j] + " " + sourceSheet + " " + sourceTable + " " + variable);
                let sheet = context.workbook.worksheets.getItem(sourceSheet);
                sheet.load("items/name");
                await context.sync();
                
                let table = sheet.tables.getItem(sourceTable);
                table.clearFilters();
                console.log(context);
                
                await context.sync();

                let filter = table.columns.getItem(variable).filter;
                filter.apply({
                    filterOn: Excel.FilterOn.values,
                    values: [v[j]]
                });
                await context.sync();

                //await clearFilters(sourceSheet, sourceTable);
                //await filterTable(sourceSheet, sourceTable, variable, j);
                let visibleRange = table.getRange().getVisibleView().load("values");
                await sheet.sync();

                let values = visibleRange.values;
                let rowCount = values.length;
                let columnCount = values[0].length;
                let worksheetDest = sourceSheet + variable + j;
                context.workbook.worksheets.getItemOrNullObject(worksheetDest).delete();
                let sheetDest = context.workbook.worksheets.add(worksheetDest);
                let range = sheetDest.getRangeByIndexes(0, 0, rowCount, columnCount);
                range.values = values;
                sheetDest.getUsedRange().format.autofitColumns();
                sheetDest.getUsedRange().format.autofitRows();

                let newTable = sheetDest.tables.add(range, true);
                newTable.name = worksheetDest;

                //await copyVisibleRange(sheet, table, sourceSheet + variable + j, context);
                await context.sync();
            }
        });
        resolve("ok");
    });
}


async function add(sourceSheet, sourceTable, variable, v) {
    await Excel.run(async (context) => {
        await Office.onReady();
        for (let j = 0; j < v.length; j++) {
            console.log(v[j] + " " + sourceSheet + " " + sourceTable + " " + variable);
            const sheet = context.workbook.worksheets.getItem(sourceSheet);
            const table = sheet.tables.getItem(sourceTable);
            table.clearFilters();
            let filter = table.columns.getItem(variable).filter;
            filter.apply({
                filterOn: Excel.FilterOn.values,
                values: [v[j]]
            });
            //await clearFilters(sourceSheet, sourceTable);
            //await filterTable(sourceSheet, sourceTable, variable, j);
            const visibleRange = table
                .getRange()
                .getVisibleView()
                .load("values");
            await context.sync();

            let values = visibleRange.values;
            let rowCount = values.length;
            let columnCount = values[0].length;
            let worksheetDest = sourceSheet + variable + j;
            context.workbook.worksheets.getItemOrNullObject(worksheetDest).delete();
            let sheetDest = context.workbook.worksheets.add(worksheetDest);
            let range = sheetDest.getRangeByIndexes(0, 0, rowCount, columnCount);
            range.values = values;
            sheetDest.getUsedRange().format.autofitColumns();
            sheetDest.getUsedRange().format.autofitRows();

            let newTable = sheetDest.tables.add(range, true);
            newTable.name = worksheetDest;

            //await copyVisibleRange(sheet, table, sourceSheet + variable + j, context);
            await context.sync();
        }
    });
}


async function filterTable(worksheet, sourceTable, variable, value) {
    const sheet = context.workbook.worksheets.getItem(worksheet);
    const table = sheet.tables.getItem(sourceTable);
    let filter = table.columns.getItem(variable).filter;
    filter.apply({
        filterOn: Excel.FilterOn.values,
        values: [value]
    });
    await context.sync();
}

async function clearFilters(worksheet, sourceTable) {
    await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getItem(worksheet);
        const table = sheet.tables.getItem(sourceTable);
        table.clearFilters();
        await context.sync();
    });
}

async function copyVisibleRange(worksheetSource, tableSource, worksheetDest) {
    return new Office.Promise(async function (resolve) {
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
    });
}


async function log(msg) {
    console.log(msg);
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
    });

}

async function deleteLastWorksheet() {
    await Excel.run(async (context) => {
        const sheets = context.workbook.worksheets;
        sheets.load("items/name");
        await context.sync();
        if (sheets.items.length > 1) {
            const lastSheet = sheets.items[sheets.items.length - 1];
            console.log(`Deleting worksheet named "${lastSheet.name}"`);
            lastSheet.delete();
            await context.sync();
        } else {
            console.log("Unable to delete the last worksheet in the workbook");    }
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
            tables.load('name, count, headers, columns');
            await context.sync();
            let allTables = [];
            for (let j in tables.items) {
                let tableheaders = tables.items[j].columns.items;
                let alltableheaders = [];
                for (let k in tableheaders) {
                    alltableheaders.push(tableheaders[k].name);
                }
                allTables.push({tablename: tables.items[j].name, categories: alltableheaders});

            }
            allSheets.push({sheetname: sheets.items[i].name, tables: allTables});
        }
        dotNetReference.invokeMethodAsync("CallbackAllWorksheets", allSheets);
    });
}