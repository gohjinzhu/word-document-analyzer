// Initialize the add-in
Office.initialize = function (reason) {
    $(document).ready(function () {
        $('#analyze').click(analyzeDocument);
        $('#runTests').click(runTests);
    });
};

function analyzeDocument() {
    Word.run(function (context) {
        var body = context.document.body;
        var paragraphs = body.paragraphs;
        context.load(paragraphs, 'text, font');
        
        return context.sync()
            .then(function () {
                if (paragraphs.items.length === 0) {
                    throw new Error("No text found in the document");
                }

                var firstParagraph = paragraphs.items[0];
                var words = firstParagraph.text.trim().split(/\s+/);
                
                if (words.length < 3) {
                    throw new Error("Document must contain at least 3 words");
                }

                // Get first word formatting
                var firstWordRange = firstParagraph.getRange('Start', 'Start').expandTo(words[0].length);
                firstWordRange.load('font');
                
                // Get second word formatting
                var secondWordStart = firstParagraph.getRange('Start', 'Start').expandTo(words[0].length + 1);
                var secondWordRange = secondWordStart.expandTo(words[1].length);
                secondWordRange.load('font');
                
                // Get third word formatting
                var thirdWordStart = secondWordStart.expandTo(words[0].length + words[1].length + 2);
                var thirdWordRange = thirdWordStart.expandTo(words[2].length);
                thirdWordRange.load('font');
                
                return context.sync()
                    .then(function () {
                        var isFirstWordBold = firstWordRange.font.bold;
                        var isSecondWordUnderlined = secondWordRange.font.underline !== 'none';
                        var thirdWordFontSize = thirdWordRange.font.size;

                        // Display results
                        var results = document.getElementById("results");
                        results.innerHTML = `
                            <h3>Analysis Results:</h3>
                            <p>First word is bold: ${isFirstWordBold}</p>
                            <p>Second word is underlined: ${isSecondWordUnderlined}</p>
                            <p>Third word font size: ${thirdWordFontSize}</p>
                        `;
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
        document: "**Bold** _Underlined_ Normal",
        expected: {
            firstWordBold: true,
            secondWordUnderlined: true,
            thirdWordFontSize: 11
        }
    },
    {
        name: "Test Case 2: No formatting",
        document: "Normal Normal Normal",
        expected: {
            firstWordBold: false,
            secondWordUnderlined: false,
            thirdWordFontSize: 11
        }
    },
    {
        name: "Test Case 3: Only bold first word",
        document: "**Bold** Normal Normal",
        expected: {
            firstWordBold: true,
            secondWordUnderlined: false,
            thirdWordFontSize: 11
        }
    },
    {
        name: "Test Case 4: Only underlined second word",
        document: "Normal _Underlined_ Normal",
        expected: {
            firstWordBold: false,
            secondWordUnderlined: true,
            thirdWordFontSize: 11
        }
    },
    {
        name: "Test Case 5: Different font size for third word",
        document: "Normal Normal **Large**",
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
    
    testCases.forEach(function (testCase) {
        Word.run(function (context) {
            // Clear document
            context.document.body.clear();
            
            // Insert test text
            var paragraph = context.document.body.insertParagraph(testCase.document, "Start");
            
            // Apply formatting based on test case
            var words = testCase.document.split(/\s+/);
            
            // First word
            if (testCase.expected.firstWordBold) {
                var firstWordRange = paragraph.getRange('Start', 'Start').expandTo(words[0].length);
                firstWordRange.font.bold = true;
            }
            
            // Second word
            if (testCase.expected.secondWordUnderlined) {
                var secondWordStart = paragraph.getRange('Start', 'Start').expandTo(words[0].length + 1);
                var secondWordRange = secondWordStart.expandTo(words[1].length);
                secondWordRange.font.underline = 'single';
            }
            
            // Third word
            if (testCase.expected.thirdWordFontSize !== 11) {
                var thirdWordStart = paragraph.getRange('Start', 'Start').expandTo(words[0].length + words[1].length + 2);
                var thirdWordRange = thirdWordStart.expandTo(words[2].length);
                thirdWordRange.font.size = testCase.expected.thirdWordFontSize;
            }
            
            return context.sync()
                .then(function () {
                    // Verify results
                    var firstWordRange = paragraph.getRange('Start', 'Start').expandTo(words[0].length);
                    firstWordRange.load('font');
                    
                    var secondWordStart = paragraph.getRange('Start', 'Start').expandTo(words[0].length + 1);
                    var secondWordRange = secondWordStart.expandTo(words[1].length);
                    secondWordRange.load('font');
                    
                    var thirdWordStart = secondWordStart.expandTo(words[0].length + words[1].length + 2);
                    var thirdWordRange = thirdWordStart.expandTo(words[2].length);
                    thirdWordRange.load('font');
                    
                    return context.sync()
                        .then(function () {
                            var actual = {
                                firstWordBold: firstWordRange.font.bold,
                                secondWordUnderlined: secondWordRange.font.underline !== 'none',
                                thirdWordFontSize: thirdWordRange.font.size
                            };
                            
                            var passed = JSON.stringify(actual) === JSON.stringify(testCase.expected);
                            
                            results.innerHTML += `
                                <div style="margin: 10px 0; padding: 10px; border: 1px solid ${passed ? 'green' : 'red'};">
                                    <h4>${testCase.name}</h4>
                                    <p>Expected: ${JSON.stringify(testCase.expected)}</p>
                                    <p>Actual: ${JSON.stringify(actual)}</p>
                                    <p>Status: ${passed ? 'PASSED' : 'FAILED'}</p>
                                </div>
                            `;
                        });
                });
        })
        .catch(function (error) {
            results.innerHTML += `
                <div style="margin: 10px 0; padding: 10px; border: 1px solid red;">
                    <h4>${testCase.name}</h4>
                    <p style="color: red;">Error: ${error.message}</p>
                </div>
            `;
        });
    });
} 