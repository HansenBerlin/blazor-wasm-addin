var GLOBAL = {};
GLOBAL.DotNetReference = null;
GLOBAL.SetDotnetReference = function (pDotNetReference) {
    GLOBAL.DotNetReference = pDotNetReference;
};

//sheet activate nach auswahl


async function wrapper() {
    await setup();
    await createWorkbook()
    await copyTableHeaders();
    await applyFilterFunction();
}


async function copyTableHeaders() {
    await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getItem("neu");
        const targetCell = sheet.getRange("A1");
        targetCell.formulas = [["=ExpensesTable[#Headers]"]];
        const spillRange = targetCell.getSpillingToRange();
        spillRange.load("address");
        sheet.getUsedRange().format.autofitColumns();
        await context.sync();
    });
}

async function applyFilterFunction() {
    await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getItem("neu");
        const targetCell = sheet.getRange("A2");
        targetCell.formulas = [
            ['=FILTER(Sample!ExpensesTable[#All], Sample!ExpensesTable[[#All],[Category]]="Groceries", "")']
        ];
        const spillRange = targetCell.getSpillingToRange();
        spillRange.load("address");
        sheet.getUsedRange().format.autofitColumns();
        await context.sync();
    });
}

async function createWorkbook() {
    await Excel.run(async (context) => {
        context.workbook.worksheets.getItemOrNullObject("neu").delete();
        const sheet = context.workbook.worksheets.add("neu");
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



async function listWorksheets() {
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
        GLOBAL.DotNetReference.invokeMethodAsync("CallbackAllWorksheets", allSheets);
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

