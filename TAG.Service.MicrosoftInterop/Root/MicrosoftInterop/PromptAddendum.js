
function ShowExcelOptions(Button)
{
	var Div = document.getElementById("ExcelDiv");
	if (Div)
	{
		if (Div.style.display !== 'none')
		{
			Div.style.display = 'none';
			Button.className = "posButton";
		}
		else
		{
			Div.style.display = 'block';
			Button.className = "posButtonPressed";
		}
	}
}

function UploadDocument()
{
	var FileInput = document.getElementById("ExcelFile");
	var File = FileInput.files[0];

	var xhttp = new XMLHttpRequest();
	xhttp.onreadystatechange = function ()
	{
		if (xhttp.readyState === 4)
		{
			if (xhttp.status === 200)
			{
				var Script = document.getElementById("script");
				Script.value = xhttp.responseText;
				FileInput.value = "";
				EvaluateExpression();
			}
			else
				ShowError(xhttp);
		}
	};

	xhttp.open("POST", "/MicrosoftInterop/ExcelToScript", true);
	xhttp.setRequestHeader("Content-Type", File.type);
	xhttp.setRequestHeader("X-FileName", File.name);
	xhttp.send(File);
}
