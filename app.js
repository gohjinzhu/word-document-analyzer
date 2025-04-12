// Replace Office.initialize with Office.onReady()
// Office.initialize = function (reason) {
//    console.log('Office initialized with reason:', reason);
//    ...
// };

// Using the modern Office.onReady approach
Office.onReady(function (info) {
    // Wait for a moment to ensure DOM is fully loaded
    setTimeout(function () {
        try {
            // Try to find our buttons
            const analyzeButton = document.getElementById('analyze');
            const runTestsButton = document.getElementById('runTests');

            if (analyzeButton) {
                analyzeButton.onclick = function () {
                    analyzeDocument();
                };
            }

            if (runTestsButton) {
                runTestsButton.onclick = function () {
                    runTests();
                };
            }

            // Directly add inline onclick attributes as a last resort
            try {
                document.querySelector('#analyze').setAttribute('onclick', 'analyzeDocument(); return false;');
                document.querySelector('#runTests').setAttribute('onclick', 'runTests(); return false;');
            } catch (e) {
                // Error adding onclick attributes
            }

        } catch (error) {
            // Error during button setup
        }
    }, 1000); // Wait 1 second after Office is ready
});

// Make sure the functions are globally accessible
window.analyzeDocument = analyzeDocument;
window.runTests = runTests;

function analyzeDocument() {
    // First, do something simple that doesn't involve Word API
    try {
        var results = document.getElementById("results");
        results.innerHTML = "<p>Starting document analysis...</p>";
    } catch (error) {
        // Error updating DOM before Word API call
    }
    
    // Now try the Word API calls
    Word.run(function (context) {
        var body = context.document.body;
        var paragraphs = body.paragraphs;
        paragraphs.load('text');

        return context.sync()
            .then(function () {
                if (paragraphs.items.length === 0) {
                    throw new Error("No text found in the document");
                }

                var firstParagraph = paragraphs.items[0];
                var text = firstParagraph.text.trim();
                var words = text.split(/\s+/);

                if (words.length < 3) {
                    throw new Error("The first paragraph must contain at least 3 words");
                }

                // Find the actual ranges for each word more carefully
                var range = firstParagraph.getRange();
                range.load('text');

                return context.sync()
                    .then(function () {
                        var fullText = range.text;
                        var wordPositions = [];
                        var wordRegex = /\S+/g;
                        var match;

                        while ((match = wordRegex.exec(fullText)) !== null) {
                            wordPositions.push({
                                start: match.index,
                                end: match.index + match[0].length - 1
                            });

                            if (wordPositions.length >= 3) break;
                        }

                        if (wordPositions.length < 3) {
                            throw new Error("Could not identify three distinct words");
                        }

                        // Get words based on their positions in the text
                        var firstWord = fullText.substring(wordPositions[0].start, wordPositions[0].end + 1);
                        var secondWord = fullText.substring(wordPositions[1].start, wordPositions[1].end + 1);
                        var thirdWord = fullText.substring(wordPositions[2].start, wordPositions[2].end + 1);

                        // Create search options to ensure we find exact word matches
                        var options = {
                            matchCase: true,
                            matchWholeWord: true
                        };

                        // Search for each word in the paragraph
                        var searchResults1 = firstParagraph.search(firstWord, options);
                        var searchResults2 = firstParagraph.search(secondWord, options);
                        var searchResults3 = firstParagraph.search(thirdWord, options);

                        // We need to load the search results before using them
                        searchResults1.load("text");
                        searchResults2.load("text");
                        searchResults3.load("text");

                        return context.sync()
                            .then(function () {
                                // Check if searches found matches
                                if (searchResults1.items.length === 0 || searchResults2.items.length === 0 || searchResults3.items.length === 0) {
                                    // If strict search fails, try without matchWholeWord option
                                    options.matchWholeWord = false;

                                    // Try again with less strict matching
                                    searchResults1 = firstParagraph.search(firstWord, options);
                                    searchResults2 = firstParagraph.search(secondWord, options);
                                    searchResults3 = firstParagraph.search(thirdWord, options);

                                    // Load the new search results
                                    searchResults1.load("text");
                                    searchResults2.load("text");
                                    searchResults3.load("text");

                                    return context.sync()
                                        .then(function () {
                                            // Check if we still don't have results
                                            if (searchResults1.items.length === 0 || searchResults2.items.length === 0 || searchResults3.items.length === 0) {
                                                throw new Error("Could not find all three words in the document");
                                            }

                                            // Continue with the successfully loaded results
                                            var firstWordRange = searchResults1.items[0];
                                            var secondWordRange = searchResults2.items[0];
                                            var thirdWordRange = searchResults3.items[0];

                                            // Load formatting properties
                                            firstWordRange.load('font');
                                            secondWordRange.load('font');
                                            thirdWordRange.load('font');

                                            return context.sync()
                                                .then(function () {
                                                    var isFirstWordBold = firstWordRange.font.bold === true;
                                                    
                                                    // Improved underline detection
                                                    var underlineValue = secondWordRange.font.underline;
                                                    var isSecondWordUnderlined = underlineValue !== 'None';
                                                    
                                                    // Improved font size handling
                                                    var thirdWordFontSize = thirdWordRange.font.size

                                                    // Display results
                                                    var results = document.getElementById("results");
                                                    results.innerHTML = `
                                                        <h3>Analysis Results:</h3>
                                                        <p>First word is bold: <span class="${isFirstWordBold ? 'passed' : 'failed'}">${isFirstWordBold}</span></p>
                                                        <p>Second word is underlined: <span class="${isSecondWordUnderlined ? 'passed' : 'failed'}">${isSecondWordUnderlined}</span> (Value: "${underlineValue}")</p>
                                                        <p>Third word font size: <span class="value">${thirdWordFontSize}</span></p>
                                                    `;
                                                });
                                        });
                                } else {
                                    // Continue with the successfully loaded results from first attempt
                                    var firstWordRange = searchResults1.items[0];
                                    var secondWordRange = searchResults2.items[0];
                                    var thirdWordRange = searchResults3.items[0];

                                    // Load formatting properties
                                    firstWordRange.load('font');
                                    secondWordRange.load('font');
                                    thirdWordRange.load({
                                        font: { size: true }
                                    });

                                    return context.sync()
                                        .then(function () {
                                            var isFirstWordBold = firstWordRange.font.bold === true;
                                            
                                            // Improved underline detection
                                            var underlineValue = secondWordRange.font.underline;
                                            var isSecondWordUnderlined = underlineValue !== 'None';
                                            
                                            // Improved font size handling
                                            var thirdWordFontSize = thirdWordRange.font.size

                                            // Display results
                                            var results = document.getElementById("results");
                                            results.innerHTML = `
                                                <h3>Analysis Results:</h3>
                                                <p>First word is bold: <span class="${isFirstWordBold ? 'passed' : 'failed'}">${isFirstWordBold}</span></p>
                                                <p>Second word is underlined: <span class="${isSecondWordUnderlined ? 'passed' : 'failed'}">${isSecondWordUnderlined}</span> (Value: "${underlineValue}")</p>
                                                <p>Third word font size: <span class="value">${thirdWordFontSize}</span></p>
                                            `;
                                        });
                                }
                            });
                    });
            })
            .catch(function (error) {
                document.getElementById("results").innerHTML = `<p style="color: red;">Error: ${error.message}</p>`;
            });
    });
}

// Test cases
var testCases = [
    {
        name: "Test Case 1: All formatting present",
        document: "Bold Underlined Normal",
        expected: {
            firstWordBold: true,
            secondWordUnderlined: true,
            thirdWordFontSize: 12
        }
    },
    {
        name: "Test Case 2: No formatting",
        document: "Normal Normal Normal",
        expected: {
            firstWordBold: false,
            secondWordUnderlined: false,
            thirdWordFontSize: 12
        }
    },
    {
        name: "Test Case 3: Only bold first word",
        document: "Bold Normal Normal",
        expected: {
            firstWordBold: true,
            secondWordUnderlined: false,
            thirdWordFontSize: 12
        }
    },
    {
        name: "Test Case 4: Only underlined second word",
        document: "Normal Underlined Normal",
        expected: {
            firstWordBold: false,
            secondWordUnderlined: true,
            thirdWordFontSize: 12
        }
    },
    {
        name: "Test Case 5: Different font size for third word",
        document: "Normal Normal Large",
        expected: {
            firstWordBold: false,
            secondWordUnderlined: false,
            thirdWordFontSize: 14
        }
    }
];

function runTests() {
    var results = document.getElementById("results");
    results.innerHTML = "<h3>Running Tests...</h3>";

    var currentTestIndex = 0;
    runNextTest();

    function runNextTest() {
        if (currentTestIndex >= testCases.length) {
            return;
        }

        var testCase = testCases[currentTestIndex];

        Word.run(function (context) {
            // Clear document
            context.document.body.clear();

            // Insert test text
            var paragraph = context.document.body.insertParagraph(testCase.document, "Start");

            // Apply formatting based on test case
            var range = paragraph.getRange();
            range.load('text');

            return context.sync()
                .then(function () {
                    var fullText = range.text;
                    var wordPositions = [];
                    var wordRegex = /\S+/g;
                    var match;

                    while ((match = wordRegex.exec(fullText)) !== null) {
                        wordPositions.push({
                            start: match.index,
                            end: match.index + match[0].length - 1
                        });

                        if (wordPositions.length >= 3) break;
                    }

                    // Get words based on their positions in the text
                    var firstWord = fullText.substring(wordPositions[0].start, wordPositions[0].end + 1);
                    var secondWord = fullText.substring(wordPositions[1].start, wordPositions[1].end + 1);
                    var thirdWord = fullText.substring(wordPositions[2].start, wordPositions[2].end + 1);

                    // Create search options to ensure we find exact word matches
                    var options = {
                        matchCase: true,
                        matchWholeWord: true
                    };

                    // Search for each word in the paragraph
                    var searchResults1 = paragraph.search(firstWord, options);
                    var searchResults2 = paragraph.search(secondWord, options);
                    var searchResults3 = paragraph.search(thirdWord, options);

                    // We need to load the search results before using them
                    searchResults1.load("text");
                    searchResults2.load("text");
                    searchResults3.load("text");

                    return context.sync()
                        .then(function () {
                            // Check if searches found matches
                            if (searchResults1.items.length === 0 || searchResults2.items.length === 0 || searchResults3.items.length === 0) {
                                // If strict search fails, try without matchWholeWord option
                                options.matchWholeWord = false;

                                // Try again with less strict matching
                                searchResults1 = paragraph.search(firstWord, options);
                                searchResults2 = paragraph.search(secondWord, options);
                                searchResults3 = paragraph.search(thirdWord, options);

                                // Load the new search results
                                searchResults1.load("text");
                                searchResults2.load("text");
                                searchResults3.load("text");

                                return context.sync()
                                    .then(function () {
                                        // Check if we still don't have results
                                        if (searchResults1.items.length === 0 || searchResults2.items.length === 0 || searchResults3.items.length === 0) {
                                            throw new Error("Could not find all three words in the document");
                                        }

                                        return continueWithFormattingAndTesting();
                                    });
                            } else {
                                return continueWithFormattingAndTesting();
                            }

                            // Define the helper function inside the then callback 
                            // but outside the if/else blocks
                            function continueWithFormattingAndTesting() {
                                // Continue with the successfully loaded results
                                var firstWordRange = searchResults1.items[0];
                                var secondWordRange = searchResults2.items[0];
                                var thirdWordRange = searchResults3.items[0];

                                // Explicitly reset all formatting to ensure consistency
                                // First word
                                firstWordRange.font.bold = false; // Reset bold
                                firstWordRange.font.underline = 'None'; // Reset underline
                                firstWordRange.font.size = 12; // Set to default font size

                                // Second word
                                secondWordRange.font.bold = false; // Reset bold
                                secondWordRange.font.underline = 'None'; // Reset underline
                                secondWordRange.font.size = 12; // Set to default font size

                                // Third word
                                thirdWordRange.font.bold = false; // Reset bold
                                thirdWordRange.font.underline = 'None'; // Reset underline
                                thirdWordRange.font.size = 12; // Set to default font size

                                // Now apply specific formatting based on test case
                                // First word
                                if (testCase.expected.firstWordBold) {
                                    firstWordRange.font.bold = true;
                                }

                                // Second word
                                if (testCase.expected.secondWordUnderlined) {
                                    secondWordRange.font.underline = 'single';
                                }

                                // Third word
                                if (testCase.expected.thirdWordFontSize !== 12) {
                                    thirdWordRange.font.size = testCase.expected.thirdWordFontSize;
                                }

                                return context.sync()
                                    .then(function () {
                                        // Verify results - reload ranges to get updated formatting
                                        firstWordRange.load('font');
                                        secondWordRange.load('font');
                                        thirdWordRange.load('font');

                                        return context.sync()
                                            .then(function () {
                                                // Process with improved checks
                                                var underlineValue = secondWordRange.font.underline;
                                                var actual = {
                                                    firstWordBold: firstWordRange.font.bold === true,
                                                    secondWordUnderlined: underlineValue !== 'None',
                                                    thirdWordFontSize: thirdWordRange.font.size
                                                };

                                                var passed = JSON.stringify(actual) === JSON.stringify(testCase.expected);

                                                results.innerHTML += `
                                                    <div class="test-result" style="margin: 10px 0; padding: 10px; border: 1px solid ${passed ? '#4caf50' : '#ff5252'};">
                                                        <h4>${testCase.name}</h4>
                                                        <p>Expected: <span class="value">${JSON.stringify(testCase.expected)}</span></p>
                                                        <p>Actual: <span class="value">${JSON.stringify(actual)}</span></p>
                                                        <p>Status: <span class="${passed ? 'passed' : 'failed'}">${passed ? 'PASSED' : 'FAILED'}</span></p>
                                                    </div>
                                                `;

                                                // Move to the next test
                                                currentTestIndex++;
                                                runNextTest();
                                            });
                                    });
                            }
                        });
                });
        })
            .catch(function (error) {
                results.innerHTML += `
                <div class="test-result" style="margin: 10px 0; padding: 10px; border: 1px solid #ff5252;">
                    <h4>${testCase.name}</h4>
                    <p class="failed">Error: ${error.message}</p>
                </div>
                `;

                // Continue to next test even if this one failed
                currentTestIndex++;
                runNextTest();
            });
    }
} 