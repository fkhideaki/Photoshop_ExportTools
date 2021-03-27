
function getDocPath(doc)
{
    try
    {
        var dp = doc.path;
        return decodeURI(dp);
    }
    catch (e)
    {
        return null;
    }
}

function getFirstValidDocPath(doc)
{
    var docPath = getDocPath(doc);
    if (docPath != null)
        return docPath;
    var nd = documents.length;
    for (var i = 0; i < nd; ++i)
    {
        var dp = getDocPath(documents[i]);
        if (dp != null)
            return dp;
    }

    return null;
}

function getBasePath(doc)
{
    var docPath = getFirstValidDocPath(doc);
    if (docPath == null)
    {
        alert("ソースパス取得に失敗");
        return null;
    }

    if (!Folder(docPath).exists)
    {
        var s = "Folder not found\n";
        s += docPath;
        alert(s);
        return null;
    }

    var name = decodeURI(doc.name).replace(/\.[^\.]+$/, '');
    var baseName = docPath + "/" + name;

    return baseName;
}

function getValidSavePath(doc, ext)
{
    var baseName = getBasePath(doc);
    if (baseName == null)
        return null;

    var savePath = baseName + ext;
    var cnt = 0;
    for (;;)
    {
        if (!File(savePath).exists)
            break;
        var i = ('000' + (++cnt)).slice(-3);
        savePath = baseName + "_" + i + ext;
    }

    return savePath;
}

function saveTIF(doc, savePath)
{
    var opt = new TiffSaveOptions;
    opt.embedColorProfile = true;

    var saveFile = File(savePath);
    doc.saveAs(saveFile, opt, true, Extension.LOWERCASE);
}

function main()
{
    if (!documents.length)
        return;

    var doc = app.activeDocument;

    var savePath = getValidSavePath(doc, ".tif");
    if (savePath == null)
        return;

    saveTIF(doc, savePath);
}

main();
