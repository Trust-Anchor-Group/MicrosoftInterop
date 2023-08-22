function UploadDocument()
{
    var FileInput = document.getElementById("WordFile");
    var File = FileInput.files[0];

    var xhttp = new XMLHttpRequest();
    xhttp.onreadystatechange = function ()
    {
        if (xhttp.readyState === 4)
        {
            if (xhttp.status === 200)
            {
                var Markdown = document.getElementById("Markdown");
                Markdown.value = xhttp.responseText;
                FileInput.value = "";
            }
            else
            {
                ShowError(xhttp);

                if (Protect)
                    document.getElementById("Password").value = "";
            }
        }
    };

    xhttp.open("POST", "//MicrosoftInterop/WordToMarkdown", true);
    xhttp.setRequestHeader("Content-Type", File.type);
    xhttp.setRequestHeader("X-FileName", File.name);
    xhttp.send(File);
}
