async function wrapper(sheetNameNew, sourceTable, variable, value) {
    await createSheet(sheetNameNew)
    await copyTableHeaders(sourceTable, sheetNameNew);
    await applyFilterFunction(sheetNameNew, sourceTable, variable, value);
    //await createTable(sheetNameNew);
}


async function write(tables, cats) {
    console.log(tables + " " + cats)
}
async function createTable(tablesheet){
    await Excel.run(async (context) => {
        let sheet = context.workbook.worksheets.getItem(tablesheet);
        let range = sheet.getUsedRange();
        range.load("address");
        await context.sync();
        let expensesTable = sheet.tables.add(range, true);
        expensesTable.name = tablesheet;
        await context.sync();
    });
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

/** Create a new table with sample data */
async function setup() {
    await Excel.run(async (context) => {
        context.workbook.worksheets.getItemOrNullObject("Sample").delete();
        const sheet = context.workbook.worksheets.add("Sample");

        const expensesTable = sheet.tables.add("A4:D4", true /*hasHeaders*/);
        expensesTable.name = "ExpensesTable";

        expensesTable.getHeaderRowRange().values = [["Date", "Merchant", "Category", "Amount"]];

        expensesTable.rows.add(null /*add at the end*/, [
            ["1/1/2020", "The Phone Company", "Communications", "$120"],
            ["1/2/2020", "Northwind Electric Cars", "Transportation", "$142"],
            ["1/5/2020", "Best For You Organics Company", "Groceries", "$27"],
            ["1/10/2020", "Coho Vineyard", "Restaurant", "$33"],
            ["1/11/2020", "Bellows College", "Education", "$350"],
            ["1/15/2020", "Trey Research", "Other", "$135"],
            ["1/15/2020", "Best For You Organics Company", "Groceries", "$97"]
        ]);

        sheet.getRange("A2:H2").values = [["Transactions", , , , , , "Category", "Groceries"]];
        sheet.getRange("A2").style = "Heading1";
        sheet.getRange("G2").style = "Heading2";
        sheet.getRange("H2").format.fill.color = "#EEEE99";

        sheet.getUsedRange().format.autofitColumns();
        sheet.getUsedRange().format.autofitRows();

        sheet.activate();
        await context.sync();
    });
}

/** Default helper for invoking an action and handling errors. */
async function tryCatch(callback) {
    try {
        await callback();
    } catch (error) {
        // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
        console.error(error);
    }
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


async function addTable(dotNetReference) {await Office.onReady();
    await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.get("neu");
        const expensesTable = sheet.tables.add("A1:D1", true /*hasHeaders*/);
        expensesTable.name = "ExpensesTable";

        expensesTable.getHeaderRowRange().values = [["Date", "Merchant", "Category", "Amount"]];

        expensesTable.rows.add(null /*add at the end*/, [
            ["1/1/2017", "The Phone Company", "Communications", "$120"],
            ["1/2/2017", "Northwind Electric Cars", "Transportation", "$142"],
            ["1/5/2017", "Best For You Organics Company", "Groceries", "$27"],
            ["1/10/2017", "Coho Vineyard", "Restaurant", "$33"],
            ["1/11/2017", "Bellows College", "Education", "$350"],
            ["1/15/2017", "Trey Research", "Other", "$135"],
            ["1/15/2017", "Best For You Organics Company", "Groceries", "$97"]
        ]);

        sheet.getUsedRange().format.autofitColumns();
        sheet.getUsedRange().format.autofitRows();

        sheet.activate();
    });
}


async function getTables2(dotNetReference) {await Office.onReady();
    await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.get("neu");
        const expensesTable = sheet.tables.get();
        expensesTable.name = "ExpensesTable";

        expensesTable.getHeaderRowRange().values = [["Date", "Merchant", "Category", "Amount"]];

        expensesTable.rows.add(null /*add at the end*/, [
            ["1/1/2017", "The Phone Company", "Communications", "$120"],
            ["1/2/2017", "Northwind Electric Cars", "Transportation", "$142"],
            ["1/5/2017", "Best For You Organics Company", "Groceries", "$27"],
            ["1/10/2017", "Coho Vineyard", "Restaurant", "$33"],
            ["1/11/2017", "Bellows College", "Education", "$350"],
            ["1/15/2017", "Trey Research", "Other", "$135"],
            ["1/15/2017", "Best For You Organics Company", "Groceries", "$97"]
        ]);

        sheet.getUsedRange().format.autofitColumns();
        sheet.getUsedRange().format.autofitRows();

        sheet.activate();
    });
}

