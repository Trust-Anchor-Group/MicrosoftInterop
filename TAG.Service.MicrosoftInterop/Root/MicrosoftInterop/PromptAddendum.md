<button type="button" onclick="ShowExcelOptions(this);">Excel</button>

<div id="ExcelDiv" style="display:none">
<p>
<label for="ExcelFile">MS Excel Document (`.xlsx`):</label>  
<input type="file" id="ExcelFile" accept="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" multiple="false"/>
</p>

<button type="button" onclick="UploadDocument()">Upload Document</button>
</div>
