const archiver = require("archiver");
const fs = require("fs");
const path = require("path");

// Copy the flows directory
function copyFlowDirectory(srcName, envType, graphUrl, loginUrl) {
    // Copy the directory
    var dstName = srcName + envType;
    fs.cpSync(srcName, dstName, { recursive: true });

    // Get the definition file
    var defFile = findDefinitionFile(dstName);
    console.log("Updating the definition file for the environment type: " + envType);
    if (defFile) {
        // Read the file
        var data = fs.readFileSync(path.join(__dirname, defFile), 'utf8');

        // Replace the graph url: graph.microsoft.com
        var content = data.replace(/graph\.microsoft\.com/g, graphUrl);

        // Replace the login url: login.microsoftonline.com
        content = content.replace(/login\.microsoftonline\.com/g, loginUrl);

        fs.writeFileSync(path.join(__dirname, defFile), content, 'utf8', function (err) {
            if (err) return console.log(err);
        });
    }
}

function findDefinitionFile(src) {
    var files = fs.readdirSync(src);
    for (var i = 0; i < files.length; i++) {
        var file = files[i];
        var path = src + "/" + file;

        // See if this is a directory
        if (fs.statSync(path).isDirectory()) {
            return findDefinitionFile(path);
        }
        // Else, see if this is the file we are looking for
        else if (file == "definition.json") {
            return path;
        }
    }
}

function generatePackage(srcDir, dstFile) {
    // Create the archive instance
    const archive = archiver("zip");
    archive.on("close", () => {
        console.log(archive.pointer() + " total bytes");
    });
    archive.on("error", (err) => {
        throw err;
    });

    // Create the zip file
    const oStream = fs.createWriteStream(dstFile);
    archive.pipe(oStream);

    // Get the files in the directory
    archive.directory("./flows/" + srcDir, false);

    // Close the file
    return archive.finalize();
}

// Read the flows directory
const srcDir = fs.readdirSync(path.join(__dirname, "flows"));
srcDir.forEach((dirName) => {
    // Determine the destination file
    const dstFile = path.join(__dirname, "dist/" + dirName + ".zip");
    const dstFileDoD = path.join(__dirname, "dist/" + dirName + "DoD.zip");

    // See if the destination files exists
    console.log("Removing the existing packages...");
    [dstFile, dstFileDoD].forEach(name => {
        if (fs.existsSync(name)) {
            // Delete the file
            fs.unlinkSync(name);

            // Log
            console.log("Deleted the flow package: " + name);
        }
    });

    // Copy the env directory
    console.log("Creating the environment directories...");
    copyFlowDirectory("./flows/" + dirName, "DoD", "graph.microsoft.us", "login.microsoftonline.us");

    // Generate the packages
    console.log("Generating the packages....");
    var promises = [];
    promises.push(generatePackage(dirName, dstFile));
    promises.push(generatePackage(dirName + "DoD", dstFileDoD));

    // Wait for the packages to be completed
    Promise.all(promises).then(() => {
        // Delete the env directory
        console.log("Removing the environment directories...");
        fs.rmSync("./flows/" + dirName + "DoD", { recursive: true, force: true });
    });
});