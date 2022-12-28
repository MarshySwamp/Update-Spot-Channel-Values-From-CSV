/*
Update Spot Channel Values - From CSV.jsx
https://github.com/MarshySwamp/Update-Spot-Channel-Values-From-CSV
v1.2 - 28th December 2022, Stephen Marsh

For background info on this script:
    https://prepression.blogspot.com/2022/12/photoshop-2022-pantone-color-book.html
    https://community.adobe.com/t5/photoshop-ecosystem-discussions/pantone-colors-converted-to-black-error/td-p/13265464
    https://community.adobe.com/t5/photoshop-ecosystem-discussions/changes-to-pantone-color-books/m-p/13353250
    https://helpx.adobe.com/creative-cloud/adobe-color.html

Notes:
You must create a CSV file titled "Photoshop Spot Channel Lab Values.csv" in your ~user account /Documents directory.
There should be a header row, then a separate row for each colour name and it's Lab values, such as:

Name,L*,a*,b*
JOKER Cola Red,60,80,90
JOKER Cola Yellow,85,3,95
etc.

* Spot colour names are CASE Sensitive!
* You will need to have a Pantone Connect Premium Subscription in order to access the Lab values for Pantone solid colours.
* Do not share/distribute your Pantone CSV lookup file.
* A log file titled "Update Spot Channel Values Log.txt" will be created on the desktop.

Thanks to @schroef for testing and feedback!
*/

#target photoshop

// Remove the old log file if it exists
var logFile = new File('~/Desktop/Update Spot Channel Values Log.txt');
if (logFile.exists)
logFile.remove();

// Validate the CSV file
var csvFile = File("~/Documents/Photoshop Spot Channel Lab Values.csv");
if (csvFile.exists) {
    if (csvFile.length > 0) {
        // Setup the script timer
        var timeDiff = {
            setStartTime: function () {
                d = new Date();
                time = d.getTime();
            },
            getDiff: function () {
                d = new Date();
                t = d.getTime() - time;
                time = d.getTime();
                return t;
            }
        };
        timeDiff.setStartTime();

        // Save the current active channels
        var savedChannelState = activeDocument.activeChannels;

        // Save the channel visibility - by Mike Bro
        var visibleChannels = visibleChannelNames();

        // Set the channel processing counter
        var processedChannelCounter = 0;

        // For loop from the late Michael_L_Hale
        for (var channelIndex = 0; channelIndex < activeDocument.channels.length; channelIndex++) {
            if (activeDocument.channels[channelIndex].kind == "ChannelType.SPOTCOLOR") {
                activeDocument.activeChannels = [activeDocument.channels[channelIndex]];
                // Channel name
                var chaName = activeDocument.activeChannels.toString().replace(/(?:\[Channel )(.+)(?:\])/, '$1');
                // Run the main function
                processSpotChannels();
            }
        }

        // Restore the saved active channels
        activeDocument.activeChannels = savedChannelState;

        // Restore the channel visibility - by Mike Bro
        for (var i = 0; i < visibleChannels.length; i++) {
            var ActiveChannel = visibleChannels[i];
            activeDocument.channels[ActiveChannel].visible = true;
        }

        // End of script notification/timer
        var spotChannelCounter = 0;
        for (var channelIndex = 0; channelIndex < activeDocument.channels.length; channelIndex++) {
            if (activeDocument.channels[channelIndex].kind == "ChannelType.SPOTCOLOR") {
                spotChannelCounter++;
            }
        }
        alert(processedChannelCounter + " of " + spotChannelCounter + " spot channels processed!" + "\n" + "(" + timeDiff.getDiff() / 1000 + " seconds)");

    } else {
        app.beep();
        alert('The "Photoshop Spot Channel Lab Values.csv" file is blank!');
        // Open the ~user's Documents directory
        Folder('~/Documents').execute();
    }

} else {
    app.beep();
    alert('The "Photoshop Spot Channel Lab Values.csv" file is missing from the Documents folder!');
    // Open the ~user's Documents directory
    Folder('~/Documents').execute();
}


////////// Functions //////////

function channelLogger(csvValue1, csvValue2, csvValue3) {
    var os = $.os.toLowerCase().indexOf("mac") >= 0 ? "mac" : "windows";
    if (os === "mac") {
        logFileLF = "Unix";
    } else {
        logFileLF = "Windows";
    }
    var logFile = new File('~/Desktop/Update Spot Channel Values Log.txt');
    var dateTime = new Date().toLocaleString();
    logFile.open("a");
    logFile.encoding = "UTF-8";
    logFile.lineFeed = logFileLF;
    logFile.writeln(dateTime);
    logFile.writeln(activeDocument.name);
    logFile.writeln(chaName);
    logFile.writeln("L: " + csvValue1 + ", a: " + csvValue2 + ", b: " + csvValue3 + "\n");
    logFile.close();
}

function processSpotChannels() {

    // CSV open, read & close
    csvFile.open("r");
    var theData = csvFile.read();
    csvFile.close();
    // CSV to array
    var theDataArray = parseCSV(theData);
    // Discard the header row
    var header = theDataArray.shift();
    // Loop over the CSV data
    for (var i = 0; i < theDataArray.length; i++) {
        // CSV array variables
        var csvName = theDataArray[i][0];
        var csvValue1 = theDataArray[i][1];
        var csvValue2 = theDataArray[i][2];
        var csvValue3 = theDataArray[i][3];
        // Case sensitive channel match to CSV name
        // To do: Is it possible to make a case insensitive regex match using a variable?
        if (csvName.match(chaName) !== null) {
            // Change spot chanel Lab values
            function s2t(s) {
                return app.stringIDToTypeID(s);
            }
            var descriptor = new ActionDescriptor();
            var descriptor2 = new ActionDescriptor();
            var descriptor3 = new ActionDescriptor();
            var reference = new ActionReference();
            reference.putEnumerated(s2t("channel"), s2t("ordinal"), s2t("targetEnum"));
            descriptor.putReference(s2t("null"), reference);
            descriptor3.putDouble(s2t("luminance"), csvValue1); // L* value
            descriptor3.putDouble(s2t("a"), csvValue2); // a* value
            descriptor3.putDouble(s2t("b"), csvValue3); // b* value
            descriptor2.putObject(s2t("color"), s2t("labColor"), descriptor3);
            descriptor.putObject(s2t("to"), s2t("spotColorChannel"), descriptor2);
            executeAction(s2t("set"), descriptor, DialogModes.NO);
            // Increment the counter for each channel processed
            processedChannelCounter++;
            // Write changes to desktop log
            channelLogger(csvValue1, csvValue2, csvValue3);
        }
    } // To do: Add logging for skipped channels?

    function parseCSV(theData, delimiter) {
        /* 
        Courtesy of William Campbell
        https://community.adobe.com/t5/photoshop-ecosystem-discussions/does-any\.$one-have-a-script-that-utilizes-a-csv-to-create-text-layers/m-p/13117458
        */

        // theData: String = contents of a CSV csvFile
        // delimiter: character that separates columns
        //            undefined defaults to comma
        // Returns: Array [[String row, String column]]

        var c = ""; // Character at index
        var d = delimiter || ","; // Default to comma
        var endIndex = theData.length;
        var index = 0;
        var maxIndex = endIndex - 1;
        var q = false; // "Are we in quotes?"
        var result = []; // Array of rows (array of column arrays)
        var row = []; // Array of columns
        var v = ""; // Column value

        while (index < endIndex) {
            c = theData[index];
            if (q) { // In quotes
                if (c == "\"") {
                    // Found quote; look ahead for another
                    if (index < maxIndex && theData[index + 1] == "\"") {
                        // Found another quote means escaped
                        // Increment and add to column value
                        index++;
                        v += c;
                    } else {
                        // Next character not a quote; last quote not escaped
                        q = !q; // Toggle "Are we in quotes?"
                    }
                } else {
                    // Add character to column value
                    v += c;
                }
            } else { // Not in quotes
                if (c == "\"") {
                    // Found quote
                    q = !q; // Toggle "Are we in quotes?"
                } else if (c == "\n" || c == "\r") {
                    // Reached end of line
                    // Test for CRLF
                    if (c == "\r" && index < maxIndex) {
                        if (theData[index + 1] == "\n") {
                            // Skip trailing newline
                            index++;
                        }
                    }
                    // Column and row complete
                    row.push(v);
                    v = "";
                    // Add row to result if first row or length matches first row
                    if (result.length === 0 || row.length == result[0].length) {
                        result.push(row);
                    }
                    row = [];
                } else if (c == d) {
                    // Found comma; column complete
                    row.push(v);
                    v = "";
                } else {
                    // Add character to column value
                    v += c;
                }
            }
            if (index == maxIndex) {
                // Reached end of theData; flush
                if (v.length || c == d) {
                    row.push(v);
                }
                // Add row to result if length matches first row
                if (row.length == result[0].length) {
                    result.push(row);
                }
                break;
            }
            index++;
        }
        return result;
    }
}

function visibleChannelNames() {
    /* 
    Courtesy of Mike Bro
    https://community.adobe.com/t5/photoshop-ecosystem-discussions/function-to-restore-original-channel-visibility/m-p/13433802
    */
    var activeChannels = [];
    var visibleChannels = activeDocument.channels;
    for (var i = 0; i < visibleChannels.length; i++) {
        if (visibleChannels[i].visible === true) {
            activeChannels.push(visibleChannels[i].name);
        }
    }
    return activeChannels;
}
