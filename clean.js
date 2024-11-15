const path = require("path");
const project = require("./package.json");
const fs = require("fs");

// Get the spfx package
const srcFile = path.join(__dirname, "dist/" + project.name + ".sppkg");

// See if it exists
if (fs.existsSync(srcFile)) {
    // Delete the file
    fs.unlinkSync(srcFile);

    // Log
    console.log("Deleted the SPFx package");
} else {
    // Log
    console.log("SPFx package doesn't exist");
}