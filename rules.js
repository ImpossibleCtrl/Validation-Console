// Utility: build expected Asset Name format
function buildExpectedAssetName(row){
    const desc = row["Asset Description"] || "";
    const wz = row["Work Zone"] || "";
    const fl = row["Floor"] || "";
    const rm = row["Room"] || "";
    const anum = row["Asset #"] || "";
    return `${desc}-${wz}-${fl}-${rm}-${anum}`.replace(/\s+/g,' ').trim();
}

function validateData(data) {
    const rowErrors = {};
    const corrected = [];
    const reportRows = [];
    const summaryCounts = {};

    // Controlled vocab lists
    const reasonNotTagged = ["Not accessible","Not safe to access","In use","Other"];
    const assetStatusVals = ["In-Service","Out-Of-Service","Stand-By","Emergency Use Only","Abandoned In Place","Seasonally In-Service","Back-Up","Removed from Facility","Critical Spare","Surplus"];
    const assetRecordStatusVals = ["Active","In-Active"];
    const caAgeVals = ["New less than 5 years old","Refurbished within 5 years","Refurbished or installed between 5 and 10 years ago","Refurbished or installed more than 10 years ago"];
    const assetConditionVals = ["5 – Excellent","4 – Good","3 – Average","2 – Poor","1 – Crisis"];
    const caEnvVals = ["Clean, temperate, dry","Wide variation in temp/humidity/dust","Extremes of temperature","Liable to extreme dust or flooding"];
    const caCapacityUnitVals = ["AMP","BTU","CFM","GAL","GPM","HP","kVA","KW","Ln.Ft.","MBH","MW","N/A","Other (List in Comments)","PSI","SCFM","Sq.Ft.","TON","V"];

    function addRowError(rowNum, msg, field){
        if(!rowErrors[rowNum]) rowErrors[rowNum] = [];
        rowErrors[rowNum].push(msg);
        const key = field+"-Error";
        summaryCounts[key] = (summaryCounts[key]||0)+1;
    }

    data.forEach((row, idx) => {
        let correctedRow = {...row};
        const rowNum = idx+2; // account for header row in Excel

        // Rule 1: Site Name required
        if(!row["Site Name"]) addRowError(rowNum,"Site Name blank","Site Name");

        // Rule 2: Workzone check
        if(row["Work Zone"] && row["Asset Name"] && !row["Asset Name"].includes(row["Work Zone"])){
            addRowError(rowNum,"Workzone mismatch","Work Zone");
        }

        // Rule 3: Building check
        if(row["Building"] && row["Asset Name"] && !row["Asset Name"].includes(row["Building"])){
            addRowError(rowNum,"Building mismatch","Building");
        }

        // Rule 4: Floor/Room check
        if(row["Floor"] && row["Asset Name"] && !row["Asset Name"].includes(row["Floor"])){
            addRowError(rowNum,"Floor mismatch","Floor");
        }
        if(row["Room"] && row["Asset Name"] && !row["Asset Name"].includes(row["Room"])){
            addRowError(rowNum,"Room mismatch","Room");
        }

        // Rule 5+6: Asset # presence
        const expectedAssetName = buildExpectedAssetName(row);
        if(!row["Asset Name"] || row["Asset Name"].trim() !== expectedAssetName){
            addRowError(rowNum,"Asset Name mismatch","Asset Name");
            correctedRow["Asset Name"] = expectedAssetName;
        }

        // Rule 7: Status Online/Offline
        if(row["Status"]){
            const val = row["Status"].trim().toLowerCase();
            if(val==="online"||val==="offline"){
                correctedRow["Status"] = val.charAt(0).toUpperCase()+val.slice(1);
            } else addRowError(rowNum,"Invalid Status","Status");
        }

        // Rule 8: Reason Not Tagged
        if(!row["TagID"] && !row["Reason Not Tagged"]){
            addRowError(rowNum,"TagID blank + Reason blank","Reason Not Tagged");
        }
        if(row["Reason Not Tagged"] && !reasonNotTagged.map(v=>v.toLowerCase()).includes(row["Reason Not Tagged"].toLowerCase())){
            addRowError(rowNum,"Invalid Reason Not Tagged","Reason Not Tagged");
        }

        // Rule 9: Asset Status
        if(row["att_Asset Status"] && !assetStatusVals.map(v=>v.toLowerCase()).includes(row["att_Asset Status"].toLowerCase())){
            addRowError(rowNum,"Invalid Asset Status","Asset Status");
        }

        // Rule 10: Asset Record Status
        if(row["att_Asset Record Status"] && !assetRecordStatusVals.map(v=>v.toLowerCase()).includes(row["att_Asset Record Status"].toLowerCase())){
            addRowError(rowNum,"Invalid Asset Record Status","Asset Record Status");
        }

        // Rule 11: In-Service Date format
        if(row["att_In-Service Date"] && !/^\d{2}\/\d{2}\/\d{4}$/.test(row["att_In-Service Date"])){
            addRowError(rowNum,"Invalid In-Service Date","In-Service Date");
        }

        // Rule 12: CA-Age
        if(row["att_In-Service Date"] && !row["att_CA-Age"]) addRowError(rowNum,"CA-Age blank","CA-Age");
        if(row["att_CA-Age"] && !caAgeVals.map(v=>v.toLowerCase()).includes(row["att_CA-Age"].toLowerCase())){
            addRowError(rowNum,"Invalid CA-Age","CA-Age");
        }

        // Rule 13: CA-Condition
        if(!row["att_CA-Condition"]) addRowError(rowNum,"CA-Condition blank","CA-Condition");

        // Rule 14: Asset Condition
        if(row["att_Asset Condition"] && !assetConditionVals.map(v=>v.toLowerCase()).includes(row["att_Asset Condition"].toLowerCase())){
            addRowError(rowNum,"Invalid Asset Condition","Asset Condition");
        }

        // Rule 15: CA-Environment
        if(row["att_CA-Environment"] && !caEnvVals.map(v=>v.toLowerCase()).includes(row["att_CA-Environment"].toLowerCase())){
            addRowError(rowNum,"Invalid CA-Environment","CA-Environment");
        }

        // Rule 16+17: Capacity
        if(row["att_Capacity Qty"] && !row["att_Capacity Unit"]){
            addRowError(rowNum,"Capacity Qty present but Unit blank","Capacity");
        }
        if(row["att_Capacity Unit"] && !caCapacityUnitVals.map(v=>v.toLowerCase()).includes(row["att_Capacity Unit"].toLowerCase())){
            addRowError(rowNum,"Invalid Capacity Unit","Capacity");
        }

        // Rule 18/26: Images
        const imgCount = ["Image","Image 2","Image 3","Image 4","Image 5"].filter(c=>row[c]).length;
        if(imgCount < 3) addRowError(rowNum,`Only ${imgCount} images provided (min 3)`,"Images");
        if(row["Asset Description"] && /(panelboard|switchgear|switchboard)/i.test(row["Asset Description"]) && imgCount < 5){
            addRowError(rowNum,"Insufficient images for Panelboard/Switchgear","Images");
        }

        // Rule 19: Manufacturer case fix
        if(row["Manufacturer"]) correctedRow["Manufacturer"] = row["Manufacturer"].charAt(0).toUpperCase()+row["Manufacturer"].slice(1).toLowerCase();

        // Rule 20: Model uppercase
        if(row["Model"]) correctedRow["Model"] = row["Model"].toUpperCase();

        // Rule 21: Serial uppercase
        if(row["Serial #"]) correctedRow["Serial #"] = row["Serial #"].toUpperCase();

        // Rule 22: JACS Code uppercase
        if(row["JACS Code"]) correctedRow["JACS Code"] = row["JACS Code"].toUpperCase();

        // Rule 23: ID uppercase
        if(row["ID"]) correctedRow["ID"] = row["ID"].toUpperCase();

        // Collect corrected + report
        corrected.push(correctedRow);
        const reportRow = {...row};
        reportRow["Row #"] = rowNum;
        reportRow["Validation Errors"] = rowErrors[rowNum] ? rowErrors[rowNum].join("; ") : "";
        reportRows.push(reportRow);
    });

    return {rowErrors,corrected,reportRows,summaryCounts};
}

function renderSummary(summaryCounts){
    const ctx = document.getElementById('summaryChart').getContext('2d');
    const labels = Object.keys(summaryCounts);
    const data = Object.values(summaryCounts);
    new Chart(ctx, {
        type: 'bar',
        data: { labels: labels, datasets: [{ label: 'Validation Issues', data: data, backgroundColor: 'rgba(255,99,132,0.6)'}] },
        options: { responsive:true, scales:{y:{beginAtZero:true}} }
    });
}

document.getElementById('validateBtn').addEventListener('click', function() {
    const fileInput = document.getElementById('fileInput');
    if (!fileInput.files.length) { alert('Please upload a CSV or XLSX file first.'); return; }
    const file = fileInput.files[0];
    const reader = new FileReader();
    reader.onload = function(e) {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, {type: 'array'});
        const sheetName = workbook.SheetNames[0];
        const ws = workbook.Sheets[sheetName];
        const jsonData = XLSX.utils.sheet_to_json(ws,{defval:""});
        const result = validateData(jsonData);

        // Show errors
        const resultsDiv = document.getElementById('results');
        resultsDiv.innerHTML = "<h3>Validation Results (Grouped by Row)</h3>";
        for(const rowNum in result.rowErrors){
            resultsDiv.innerHTML += `<p><b>Row ${rowNum}</b></p><ul>`+ result.rowErrors[rowNum].map(e=>"<li>"+e+"</li>").join("")+ "</ul>";
        }
        renderSummary(result.summaryCounts);

        // Corrected export
        const newSheet = XLSX.utils.json_to_sheet(result.corrected);
        const newWB = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(newWB, newSheet, "Corrected");
        const wbout = XLSX.write(newWB,{bookType:'xlsx',type:'array'});
        const blob = new Blob([wbout], {type:"application/octet-stream"});
        const url = URL.createObjectURL(blob);
        const dl = document.getElementById('downloadCorrected');
        dl.href = url; dl.style.display="inline-block"; dl.download = "Corrected.xlsx";

        // Validation report
        const reportSheet = XLSX.utils.json_to_sheet(result.reportRows);
        const reportWB = XLSX.utils.book_new();
        XLSX.utils.book_append_sheet(reportWB, reportSheet, "Validation Report");
        const wbout2 = XLSX.write(reportWB,{bookType:'xlsx',type:'array'});
        const blob2 = new Blob([wbout2], {type:"application/octet-stream"});
        const url2 = URL.createObjectURL(blob2);
        const dl2 = document.getElementById('downloadReport');
        dl2.href = url2; dl2.style.display="inline-block"; dl2.download = "Validation_Report.xlsx";
    };
    reader.readAsArrayBuffer(file);
});
