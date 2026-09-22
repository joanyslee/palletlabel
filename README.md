<!DOCTYPE html>
<html lang="en">
<head>
<meta charset="UTF-8">
<meta name="viewport" content="width=device-width, initial-scale=1.0">
<title>Pallet Label Generator</title>

<!-- Excel reader -->
<script src="https://cdn.jsdelivr.net/npm/xlsx@0.18.5/dist/xlsx.full.min.js"></script>

<style>
    * {
        box-sizing: border-box;
    }

    body {
        margin: 0;
        font-family: Arial, Helvetica, sans-serif;
        background: #f3f4f6;
        color: #111827;
    }

    .container {
        max-width: 1100px;
        margin: 40px auto;
        padding: 20px;
    }

    h1 {
        text-align: center;
        margin-bottom: 30px;
    }

    .upload-box {
        background: white;
        border-radius: 12px;
        padding: 25px;
        box-shadow: 0 2px 10px rgba(0,0,0,.08);
        margin-bottom: 25px;
        text-align: center;
    }

    input[type="file"] {
        margin: 15px 0;
    }

    button {
        border: none;
        background: #2563eb;
        color: white;
        padding: 12px 22px;
        border-radius: 7px;
        font-size: 16px;
        cursor: pointer;
        margin: 5px;
    }

    button:hover {
        background: #1d4ed8;
    }

    button.secondary {
        background: #374151;
    }

    button.secondary:hover {
        background: #1f2937;
    }

    #status {
        margin-top: 10px;
        color: #4b5563;
    }

    #preview {
        display: flex;
        flex-direction: column;
        align-items: center;
        gap: 30px;
    }

    /* Letter paper */
    .label {
        width: 8.5in;
        height: 11in;
        background: white;
        padding: 0.55in;
        position: relative;
        page-break-after: always;
        box-shadow: 0 2px 12px rgba(0,0,0,.15);

        display: flex;
        flex-direction: column;
    }

    .item-code {
        text-align: center;
        font-size: 42px;
        font-weight: bold;
        word-break: break-word;
        margin-top: 0.25in;
    }

    .pallet-section {
        flex: 1;
        display: flex;
        flex-direction: column;
        justify-content: center;
        align-items: center;
    }

    .pallet-number {
        font-size: 110px;
        font-weight: bold;
        line-height: 1;
        text-align: center;
        word-break: break-word;
    }

    .pallet-title {
        font-size: 28px;
        font-weight: bold;
        margin-top: 20px;
        letter-spacing: 2px;
    }

    .bottom-info {
        font-size: 30px;
        line-height: 1.8;
        margin-bottom: 0.25in;
    }

    .lot-line {
        border-bottom: 2px solid #111;
        display: inline-block;
        width: 5in;
        height: 35px;
        vertical-align: bottom;
    }

    .qty {
        font-weight: bold;
    }

    .hidden {
        display: none;
    }

    @media print {
        body {
            background: white;
        }

        .container > *:not(#preview) {
            display: none !important;
        }

        #preview {
            display: block;
        }

        .label {
            box-shadow: none;
            margin: 0;
        }
    }

    @page {
        size: Letter;
        margin: 0;
    }
</style>
</head>

<body>

<div class="container">

    <div id="controls">

        <h1>Pallet Label Generator</h1>

        <div class="upload-box">

            <h2>Upload Packing List</h2>

            <p>
                Upload your Excel (.xlsx) packing list.
            </p>

            <input
                type="file"
                id="excelFile"
                accept=".xlsx,.xls"
            >

            <br>

            <button onclick="generateLabels()">
                Generate Labels
            </button>

            <button
                class="secondary"
                onclick="window.print()"
            >
                Print / Save PDF
            </button>

            <div id="status"></div>

        </div>

    </div>

    <div id="preview"></div>

</div>


<script>

function normalizeHeader(value) {
    return String(value || "")
        .trim()
        .toLowerCase()
        .replace(/[_-]/g, " ")
        .replace(/\s+/g, " ");
}


function findColumn(headers, possibleNames) {

    for (const header of headers) {

        const normalized = normalizeHeader(header);

        for (const name of possibleNames) {

            if (normalized === normalizeHeader(name)) {
                return header;
            }

        }

    }

    return null;
}


async function generateLabels() {

    const fileInput = document.getElementById("excelFile");
    const status = document.getElementById("status");
    const preview = document.getElementById("preview");

    if (!fileInput.files.length) {
        alert("Please select an Excel file first.");
        return;
    }

    const file = fileInput.files[0];

    status.textContent = "Reading Excel file...";
    preview.innerHTML = "";

    try {

        const data = await file.arrayBuffer();

        const workbook = XLSX.read(data, {
            type: "array"
        });

        /*
         * Use the first worksheet.
         */
        const sheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[sheetName];

        const rows = XLSX.utils.sheet_to_json(
            worksheet,
            {
                defval: ""
            }
        );

        if (!rows.length) {
            throw new Error("No data found in the Excel sheet.");
        }


        const headers = Object.keys(rows[0]);


        /*
         * These are the column names the program will recognize.
         *
         * You can add more names here later if your real
         * packing list uses different column names.
         */

        const itemColumn = findColumn(headers, [
            "Item Code",
            "ItemCode",
            "Item",
            "SKU",
            "Product Code"
        ]);

        const palletColumn = findColumn(headers, [
            "Pallet Number",
            "PalletNumber",
            "Pallet",
            "Pallet No",
            "Pallet #",
            "Pallet ID"
        ]);

        const qtyColumn = findColumn(headers, [
            "Qty",
            "Quantity",
            "QTY."
        ]);


        if (!itemColumn) {
            throw new Error(
                "Could not find Item Code column."
            );
        }

        if (!palletColumn) {
            throw new Error(
                "Could not find Pallet Number column."
            );
        }

        if (!qtyColumn) {
            throw new Error(
                "Could not find Quantity / Qty column."
            );
        }


        let generated = 0;


        rows.forEach((row) => {

            const itemCode = String(
                row[itemColumn] ?? ""
            ).trim();

            const palletNumber = String(
                row[palletColumn] ?? ""
            ).trim();

            const qty = String(
                row[qtyColumn] ?? ""
            ).trim();


            /*
             * Skip completely empty rows.
             */
            if (!itemCode && !palletNumber && !qty) {
                return;
            }


            const label = document.createElement("div");

            label.className = "label";


            /*
             * ITEM CODE
             */
            const item = document.createElement("div");

            item.className = "item-code";
            item.textContent = itemCode;


            /*
             * PALLET NUMBER
             */
            const palletSection = document.createElement("div");

            palletSection.className = "pallet-section";


            const pallet = document.createElement("div");

            pallet.className = "pallet-number";
            pallet.textContent = palletNumber;


            const title = document.createElement("div");

            title.className = "pallet-title";
            title.textContent = "PALLET NUMBER";


            palletSection.appendChild(pallet);
            palletSection.appendChild(title);


            /*
             * BOTTOM INFO
             */
            const bottom = document.createElement("div");

            bottom.className = "bottom-info";

            bottom.innerHTML = `
                <div>
                    Lot Number:
                    <span class="lot-line"></span>
                </div>

                <div class="qty">
                    Qty: ${escapeHtml(qty)}
                </div>
            `;


            label.appendChild(item);
            label.appendChild(palletSection);
            label.appendChild(bottom);

            preview.appendChild(label);

            generated++;

        });


        status.textContent =
            `${generated} pallet label(s) generated.`;

    }
    catch (error) {

        console.error(error);

        status.textContent = "";

        alert(
            "Error: " + error.message
        );

    }

}


/*
 * Prevent Excel cell contents from being interpreted
 * as HTML.
 */
function escapeHtml(value) {

    return String(value)
        .replace(/&/g, "&amp;")
        .replace(/</g, "&lt;")
        .replace(/>/g, "&gt;")
        .replace(/"/g, "&quot;")
        .replace(/'/g, "&#039;");
}

</script>

</body>
</html>
