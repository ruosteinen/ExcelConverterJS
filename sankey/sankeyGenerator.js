var globalSankeyData, globalLayout;

document.getElementById('generateSankey').addEventListener('click', function() {
    const file = document.getElementById('uploadFile').files[0];
    const sheetName = document.getElementById('sheetName').value;
    const sourceCol = parseInt(document.getElementById('sourceCol').value, 10) - 1;
    const targetCol = parseInt(document.getElementById('targetCol').value, 10) - 1;
    const valueCol = parseInt(document.getElementById('valueCol').value, 10) - 1;
    const percentCol = parseInt(document.getElementById('percentCol').value, 10) - 1;

    if (!file || !sheetName || isNaN(sourceCol) || isNaN(targetCol) || isNaN(valueCol) || isNaN(percentCol)) {
        alert("Please fill all fields correctly.");
        return;
    }

    const reader = new FileReader();
    reader.onload = function(e) {
        const data = e.target.result;
        const workbook = XLSX.read(data, {type: 'binary'});
        const worksheet = workbook.Sheets[sheetName];
        if (!worksheet) {
            alert("Sheet not found.");
            return;
        }
        const json = XLSX.utils.sheet_to_json(worksheet, {header: 1});
        generateSankeyDiagram(json, sourceCol, targetCol, valueCol, percentCol);
    };
    reader.readAsBinaryString(file);
});

function generateSankeyDiagram(data, sourceCol, targetCol, valueCol, percentCol) {
    const nodesArray = [];
    const links = [];

    data.slice(1).forEach(row => {
        const source = row[sourceCol];
        const target = row[targetCol];
        const value = parseFloat(row[valueCol]);
        const percent = Math.floor(parseFloat(row[percentCol]) * 100);

        if (source && target && !isNaN(value) && !isNaN(percent)) {
            if (!nodesArray.includes(source)) nodesArray.push(source);
            if (!nodesArray.includes(target)) nodesArray.push(target);

            links.push({
                source: nodesArray.indexOf(source),
                target: nodesArray.indexOf(target),
                value: value,
                percent: percent
            });
        }
    });

    globalSankeyData = [{
        type: "sankey",
        orientation: "h",
        node: {
            pad: 15,
            thickness: 20,
            line: { color: "black", width: 0.5 },
            label: nodesArray
        },
        link: {
            source: links.map(link => link.source),
            target: links.map(link => link.target),
            value: links.map(link => link.value),
            customdata: links.map(link => [link.percent]),
            hovertemplate: 'Value: %{value}<br>Percent: %{customdata[0]}%<extra></extra>'
        }
    }];

    globalLayout = { font: { size: 10 } };

    Plotly.newPlot('sankeyDiagram', globalSankeyData, globalLayout);
}

document.getElementById('saveSankey').addEventListener('click', function() {
    const htmlContent = `
<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <title>Sankey Diagram from Excel</title>
    <script src="https://cdn.plot.ly/plotly-latest.min.js"></script>
</head>
<body>
    <div id="sankeyDiagram" style="width:100%; height:600px;"></div>
    <script>
        var data = ${JSON.stringify(globalSankeyData)};
        var layout = ${JSON.stringify(globalLayout)};
        Plotly.newPlot('sankeyDiagram', data, layout);
    </script>
</body>
</html>
`;

    const blob = new Blob([htmlContent], {type: 'text/html'});
    const url = URL.createObjectURL(blob);
    const link = document.createElement('a');
    link.href = url;
    link.download = 'sankey_diagram.html';
    document.body.appendChild(link);
    link.click();
    document.body.removeChild(link);
    URL.revokeObjectURL(url);
});
