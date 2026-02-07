/**
  * BookCreator
  * 
  * BookCreator is an advanced script for Adobe InDesign 
  * designed to automate the entire book creation process 
  * from templates. It can automatically import Markdown 
  * text and inject it into your generated book, along with variables.
  * 
  * @version 1.0 beta 3
  * @license AGPL
  * @author entremonde / Spectral lab
  * @website http://lab.spectral.art
  */

#target indesign

/**
 * @namespace BookCreator
 * @description Main module for automated book generation in InDesign
 */
var BookCreator = (function() {
    /** @constant {string} VERSION Current version of the script */
    var VERSION = "1.0";
    
    /**
     * @namespace Utils
     * @description ES3-compatible general utility functions
     * @private
     */
    
    /**
     * Checks if an array contains a specific item
     * @param {Array} array - The array to check
     * @param {*} item - The item to search for
     * @return {boolean} True if item is found, false otherwise
     */
    function arrayContains(array, item) {
        for (var i = 0; i < array.length; i++) {
            if (array[i] === item) {
                return true;
            }
        }
        return false;
    }
    
    /**
     * ExtendScript-compatible string trim function
     * @param {string} str - String to trim
     * @return {string} Trimmed string
     */
    function trim(str) {
        return str.replace(/^\s+|\s+$/g, '');
    }
    
    /**
     * Checks if an object is an array
     * @param {*} obj - Object to check
     * @return {boolean} True if the object is an array
     */
    function isArray(obj) {
        return Object.prototype.toString.call(obj) === '[object Array]';
    }
    
    /**
     * Normalizes a book title for use as filename prefix
     * @param {string} title - Book title to normalize
     * @return {string} Normalized prefix (first 1-3 words, max 8 chars, uppercase)
     */
    function normalizeBookTitle(title) {
        if (!title) return "";
        
        var words = title.split(' ');
        var result = "";
        
        // Take first 1-3 words
        for (var i = 0; i < Math.min(3, words.length); i++) {
            var word = words[i];
            // Remove ALL types of punctuation including all apostrophes
            var cleanWord = word.replace(/[?!.,;:\-_\(\)\[\]{}\u0027\u2019\u2018\u201B]/g, '');
            
            // Check if adding this word would exceed 8 characters
            var testResult = result + cleanWord;
            if (testResult.length <= 8) {
                result = testResult;
            } else {
                break;
            }
        }
        
        return result.toUpperCase();
    }
    
    /**
     * Progress bar window for long operations
     * @constructor
     * @param {string} title - Window title
     * @param {number} maxValue - Maximum value for progress
     */
    function ProgressBar(title, maxValue) {
        // Utiliser "palette" pour une fenêtre non-bloquante qui reste au premier plan
        this.window = new Window("palette", title || "Processing...");
        this.window.maximized = false;
        this.window.minimized = false;
        this.window.orientation = "column";
        this.window.alignChildren = "fill";
        this.window.preferredSize.width = 400;
        
        // Status text
        this.statusText = this.window.add("statictext", undefined, "Starting...");
        
        this.cancelled = false;
        
        this.show = function() {
            this.window.show();
        };
        
        this.close = function() {
            this.window.close();
        };
        
        this.update = function(value, status, details) {
            if (status) this.statusText.text = status;
            
            // Force le rafraîchissement multiple
            this.window.update();
            
            // Ramener au premier plan
            this.bringToFront();
            
            // Petite pause pour permettre le rafraîchissement
            app.doScript(function(){}, ScriptLanguage.JAVASCRIPT, [], UndoModes.FAST_ENTIRE_SCRIPT);
        };
        
        this.isCancelled = function() {
            return this.cancelled;
        };
        
        this.bringToFront = function() {
            try {
                this.window.active = true;
                this.window.show();
            } catch(e) {
                // Ignore errors
            }
        };
    }
    
    /**
     * @namespace YAMLParser
     * @description Module for parsing and stringifying YAML data
     */
    var YAMLParser = (function() {
        /**
         * Utility functions for YAML parsing operations
         * @type {Object}
         */
        var utils = {
            /**
             * Trims whitespace from a string
             * @param {string} str - String to trim
             * @return {string} Trimmed string
             */
            trim: function(str) {
                return str.replace(/^\s+|\s+$/g, '');
            },
            
            /**
             * Checks if an object is an array
             * @param {*} obj - Object to check
             * @return {boolean} True if object is an array
             */
            isArray: function(obj) {
                return Object.prototype.toString.call(obj) === '[object Array]';
            },
            
            /**
             * Converts YAML string values to appropriate JavaScript types
             * @param {string} value - YAML value to convert
             * @return {*} Converted value in appropriate type
             */
            convertValue: function(value) {
                // Empty value
                if (value === undefined || value === null || utils.trim(value) === "") {
                    return "";
                }
                // Booleans
                else if (value === "true" || value === "yes" || value === "on") {
                    return true;
                }
                else if (value === "false" || value === "no" || value === "off") {
                    return false;
                }
                // Null value
                else if (value === "null" || value === "~") {
                    return null;
                }
                // Numbers (avoid treating strings like "0123" as numbers)
                else if (!isNaN(parseFloat(value)) && utils.trim(value) !== "" && 
                        !/^0[0-9]+/.test(value)) {
                    return parseFloat(value);
                }
                // Quoted strings (single or double quotes)
                else if ((value.charAt(0) === '"' && value.charAt(value.length - 1) === '"') ||
                         (value.charAt(0) === "'" && value.charAt(value.length - 1) === "'")) {
                    return value.substr(1, value.length - 2);
                }
                // Inline array
                else if (value.charAt(0) === '[' && value.charAt(value.length - 1) === ']') {
                    return value.substr(1, value.length - 2).split(',').map(function(item) {
                        return utils.convertValue(utils.trim(item));
                    });
                }
                // Simple string
                else {
                    return value;
                }
            },
            
            /**
             * Checks if a string needs quotes in YAML format
             * @param {string} str - String to check
             * @return {boolean} True if quotes are needed
             */
            needsQuotes: function(str) {
                return !isNaN(parseFloat(str)) || 
                       str.match(/^(true|false|yes|no|on|off|null)$/i) ||
                       str.match(/[:#{}[\],&*!|>'"%@`]/);
            },
            
            /**
             * Escapes special characters in a string
             * @param {string} str - String to escape
             * @return {string} Escaped string
             */
            escapeString: function(str) {
                return str.replace(/"/g, '\\"').replace(/\n/g, '\\n');
            },
            
            /**
             * Removes YAML front matter delimiters
             * @param {string} str - String with possible delimiters
             * @return {string} String without delimiters
             */
            removeFrontMatterDelimiters: function(str) {
                return str.replace(/^---\s*\n/, '').replace(/\n---\s*$/, '');
            }
        };
        
        /**
         * Context for YAML parsing, maintains state during parsing
         * @constructor
         * @param {string} yamlString - YAML content to parse
         */
        function ParseContext(yamlString) {
            this.lines = utils.removeFrontMatterDelimiters(yamlString).split(/\r\n|\r|\n/);
            this.result = {};
            this.currentKey = null;
            this.currentIndent = 0;
            this.inMultiline = false;
            this.multilineValue = "";
            this.literalBlock = false;
            this.inList = false;
            this.currentList = [];
            this.inObject = false;
            this.currentObject = {};
            this.lineIndex = 0;
            this.inComplexList = false;
            this.currentComplexItem = null;
            this.complexItemIndent = 0; 
            
            /**
             * Gets the current line being processed
             * @return {string|null} Current line or null if at end
             */
            this.getCurrentLine = function() {
                return this.lineIndex < this.lines.length ? this.lines[this.lineIndex] : null;
            };
            
            /**
             * Advances to the next line and returns it
             * @return {string|null} Next line or null if at end
             */
            this.getNextLine = function() {
                this.lineIndex++;
                return this.getCurrentLine();
            };
            
            /**
             * Gets the indentation level of a line
             * @param {string} line - Line to check
             * @return {number} Number of whitespace characters
             */
            this.getLineIndent = function(line) {
                var match = line.match(/^(\s+)/);
                return match ? match[0].length : 0;
            };
            
            /**
             * Checks if a line is empty or a comment
             * @param {string} line - Line to check
             * @return {boolean} True if line is empty or a comment
             */
            this.isEmptyOrComment = function(line) {
                return !line || utils.trim(line) === "" || utils.trim(line).charAt(0) === "#";
            };
            
            /**
             * Checks if a line is a list item
             * @param {string} line - Line to check
             * @return {boolean} True if line is a list item
             */
            this.isListItem = function(line) {
                return line && line.match(/^\s*- /);
            };
            
            /**
             * Extracts key-value pair from a line
             * @param {string} line - Line to process
             * @return {Object|null} Object with key and value, or null
             */
            this.getKeyValuePair = function(line) {
                var pair = line.match(/^([^:]+):\s*(.*?)$/);
                if (pair) {
                    return {
                        key: utils.trim(pair[1]),
                        value: utils.trim(pair[2])
                    };
                }
                return null;
            };
            
            /**
             * Extracts and converts a list item value
             * @param {string} line - Line containing list item
             * @return {*} Converted list item value
             */
            this.extractListItem = function(line) {
                var item = line.replace(/^\s*- /, "");
                return utils.convertValue(utils.trim(item));
            };
            
            /**
             * Checks if a line starts a complex object (has key: value after list item)
             * @param {string} line - Line to check
             * @return {boolean} True if line starts a complex object
             */
            this.isComplexListItem = function(line) {
                // Check if it's a list item followed by a key:value on the same line
                var match = line.match(/^\s*-\s+([^:]+):\s*(.*?)$/);
                return match !== null;
            };
            
            /**
             * Extracts key-value from a complex list item
             * @param {string} line - Line containing complex list item
             * @return {Object} Object with key and value
             */
            this.extractComplexListItem = function(line) {
                var match = line.match(/^\s*-\s+([^:]+):\s*(.*?)$/);
                if (match) {
                    return {
                        key: utils.trim(match[1]),
                        value: utils.trim(match[2])
                    };
                }
                return null;
            };
        }
        
        /**
         * Parses a YAML string into a JavaScript object
         * @param {string} yamlString - YAML content to parse
         * @return {Object} Parsed JavaScript object
         */
        function parse(yamlString) {
            if (!yamlString) return {};
            
            var ctx = new ParseContext(yamlString);
            var line = ctx.getCurrentLine();
            
            while (line !== null) {
                // Skip empty lines and comments
                if (ctx.isEmptyOrComment(line)) {
                    line = ctx.getNextLine();
                    continue;
                }
                
                var lineIndent = ctx.getLineIndent(line);
                
                // Handle lists
                if (ctx.inList) {
                    if (ctx.isListItem(line)) {
                        var trimmedLine = line.replace(/^\s*-\s*/, "");
                        
                        // Check if it's a complex item (has key: value)
                        if (trimmedLine.match(/^[^:]+:\s*.+$/)) {
                            // Complex list item
                            var complexPair = ctx.getKeyValuePair(trimmedLine);
                            if (complexPair) {
                                var complexObj = {};
                                complexObj[complexPair.key] = utils.convertValue(complexPair.value);
                                
                                // Look ahead for more properties of this object
                                var nextIndex = ctx.lineIndex + 1;
                                while (nextIndex < ctx.lines.length) {
                                    var nextLine = ctx.lines[nextIndex];
                                    var nextIndent = ctx.getLineIndent(nextLine);
                                    
                                    if (!ctx.isListItem(nextLine) && nextIndent > lineIndent) {
                                        var nextPair = ctx.getKeyValuePair(nextLine);
                                        if (nextPair) {
                                            complexObj[nextPair.key] = utils.convertValue(nextPair.value);
                                        }
                                        nextIndex++;
                                    } else {
                                        break;
                                    }
                                }
                                
                                ctx.currentList.push(complexObj);
                                ctx.lineIndex = nextIndex - 1;
                                line = ctx.getNextLine();
                                continue;
                            }
                        } else {
                            // Simple list item
                            ctx.currentList.push(utils.convertValue(trimmedLine));
                        }
                        
                        line = ctx.getNextLine();
                        continue;
                    } else if (lineIndent <= ctx.currentIndent) {
                        // End of list
                        ctx.result[ctx.currentKey] = ctx.currentList;
                        ctx.inList = false;
                        ctx.currentList = [];
                        ctx.currentKey = null;
                        // Don't advance, process current line
                    }
                }
                
                // Handle complex list items (objects within lists)
                if (ctx.inComplexList) {
                    if (lineIndent > ctx.complexItemIndent) {
                        // Still within the complex item
                        var itemPair = ctx.getKeyValuePair(line);
                        if (itemPair) {
                            ctx.currentComplexItem[itemPair.key] = utils.convertValue(itemPair.value);
                        }
                        line = ctx.getNextLine();
                        continue;
                    } else {
                        // End of complex item
                        ctx.currentList.push(ctx.currentComplexItem);
                        ctx.inComplexList = false;
                        ctx.currentComplexItem = null;
                        // Don't advance to next line, process current line
                    }
                }
                
                // Handle nested objects
                if (ctx.inObject) {
                    if (lineIndent > ctx.currentIndent || ctx.isListItem(line)) {
                        // Continue with object
                        var objPair = ctx.getKeyValuePair(line);
                        if (objPair) {
                            ctx.currentObject[objPair.key] = utils.convertValue(objPair.value);
                        }
                        line = ctx.getNextLine();
                        continue;
                    } else {
                        // End of nested object
                        ctx.result[ctx.currentKey] = ctx.currentObject;
                        ctx.inObject = false;
                        ctx.currentObject = {};
                        ctx.currentKey = null;
                        // Don't advance to next line
                    }
                }
                
                // Handle multiline values
                if (ctx.inMultiline) {
                    if (line.match(/^\s{2,}/) && utils.trim(line) !== "") {
                        // Still in multiline value
                        var lineContent = utils.trim(line.replace(/^\s{2,}/, ""));
                        if (ctx.literalBlock) {
                            // For literal blocks (|), preserve line breaks
                            ctx.multilineValue += (ctx.multilineValue ? "\r" : "") + lineContent;
                        } else {
                            // For folded blocks (>), join with spaces
                            ctx.multilineValue += (ctx.multilineValue ? " " : "") + lineContent;
                        }
                        line = ctx.getNextLine();
                        continue;
                    } else {
                        // End of multiline value
                        ctx.result[ctx.currentKey] = ctx.multilineValue;
                        ctx.inMultiline = false;
                        ctx.literalBlock = false;
                        ctx.currentKey = null;
                        ctx.multilineValue = "";
                        // Don't advance to next line
                    }
                }
                
                // Check for new key/value pair
                var pair = ctx.getKeyValuePair(line);
                if (pair) {
                    ctx.currentKey = pair.key;
                    var value = pair.value;
                    
                    // Check for start of list
                    if (value === "" && 
                        ctx.lineIndex + 1 < ctx.lines.length && 
                        ctx.isListItem(ctx.lines[ctx.lineIndex + 1])) {
                        ctx.inList = true;
                        ctx.currentList = [];
                        ctx.currentIndent = lineIndent;
                        line = ctx.getNextLine();
                        continue;
                    }
                    
                    // Check for start of nested object
                    if (value === "" && 
                        ctx.lineIndex + 1 < ctx.lines.length && 
                        ctx.getLineIndent(ctx.lines[ctx.lineIndex + 1]) > lineIndent && 
                        !ctx.isListItem(ctx.lines[ctx.lineIndex + 1])) {
                        ctx.inObject = true;
                        ctx.currentObject = {};
                        ctx.currentIndent = lineIndent;
                        line = ctx.getNextLine();
                        continue;
                    }
                    
                    // Check value types
                    if (value === "" || value === "|" || value === ">") {
                        // Multiline value starts on next line
                        ctx.inMultiline = true;
                        ctx.multilineValue = "";
                        ctx.literalBlock = (value === "|"); // Track if it's a literal block
                    } else {
                        // Simple value
                        ctx.result[ctx.currentKey] = utils.convertValue(value);
                    }
                    
                    line = ctx.getNextLine();
                    continue;
                }
                
                // If we reach here, we couldn't process the line
                // Move to next line to avoid infinite loop
                line = ctx.getNextLine();
            }
            
            // Final processing of in-progress structures
            if (ctx.inMultiline && ctx.currentKey) {
                ctx.result[ctx.currentKey] = ctx.multilineValue;
                ctx.literalBlock = false;
            }
            
            if (ctx.inComplexList && ctx.currentKey) {
                // Finalize current complex item and add to list
                if (ctx.currentComplexItem) {
                    ctx.currentList.push(ctx.currentComplexItem);
                }
                ctx.result[ctx.currentKey] = ctx.currentList;
                ctx.inComplexList = false;
                ctx.currentComplexItem = null;
            }
            
            if (ctx.inList && ctx.currentKey) {
                ctx.result[ctx.currentKey] = ctx.currentList;
            }
            
            if (ctx.inObject && ctx.currentKey) {
                ctx.result[ctx.currentKey] = ctx.currentObject;
            }
            
            return ctx.result;
        }
        
        /**
         * Converts a JavaScript object to YAML string
         * @param {Object} obj - JavaScript object to convert
         * @return {string} YAML formatted string
         */
        function stringify(obj) {
            if (!obj) return "";
            
            var result = "";
            
            for (var key in obj) {
                if (obj.hasOwnProperty(key)) {
                    var value = obj[key];
                    
                    if (value === null || value === undefined) {
                        result += key + ": null\n";
                    } else if (typeof value === "boolean" || typeof value === "number") {
                        result += key + ": " + value + "\n";
                    } else if (utils.isArray(value)) {
                        if (value.length === 0) {
                            result += key + ": []\n";
                        } else if (value.length <= 5 && value.every(function(item) { 
                            return typeof item === "string" && item.length < 20 && !item.includes(","); 
                        })) {
                            // Short list, inline format
                            result += key + ": [" + value.join(", ") + "]\n";
                        } else {
                            // Long or complex list, multiline format
                            result += key + ":\n";
                            for (var i = 0; i < value.length; i++) {
                                var item = value[i];
                                if (typeof item === "string" && item.includes("\n")) {
                                    // Multiline item
                                    result += "  - |\n";
                                    var lines = item.split(/\r\n|\r|\n/);
                                    for (var j = 0; j < lines.length; j++) {
                                        result += "    " + lines[j] + "\n";
                                    }
                                } else if (typeof item === "object" && item !== null) {
                                    // Complex object item
                                    result += "  -\n";
                                    for (var prop in item) {
                                        if (item.hasOwnProperty(prop)) {
                                            result += "    " + prop + ": " + stringifyValue(item[prop], 4) + "\n";
                                        }
                                    }
                                } else {
                                    // Simple item
                                    result += "  - " + stringifyValue(item, 4) + "\n";
                                }
                            }
                        }
                    } else if (typeof value === "string") {
                        if (value.indexOf("\n") !== -1) {
                            // Multiline value with pipe literal style to preserve line breaks
                            result += key + ": |\n";
                            var lines = value.split(/\r\n|\r|\n/);
                            for (var i = 0; i < lines.length; i++) {
                                result += "  " + lines[i] + "\n";
                            }
                        } else if (value === "") {
                            result += key + ":\n";
                        } else if (utils.needsQuotes(value)) {
                            result += key + ": \"" + utils.escapeString(value) + "\"\n";
                        } else {
                            result += key + ": " + value + "\n";
                        }
                    } else if (typeof value === "object") {
                        // Nested objects
                        result += key + ":\n";
                        for (var prop in value) {
                            if (value.hasOwnProperty(prop)) {
                                // Deeper indentation for nested properties
                                result += "  " + prop + ": " + stringifyValue(value[prop], 2) + "\n";
                            }
                        }
                    }
                }
            }
            
            return result;
        }
        
        /**
         * Helper function for stringifying complex values with indentation
         * @param {*} value - Value to stringify
         * @param {number} indent - Indentation level
         * @return {string} String representation
         * @private
         */
        function stringifyValue(value, indent) {
            var indentStr = Array(indent + 1).join(" ");
            
            if (value === null || value === undefined) {
                return "null";
            } else if (typeof value === "boolean" || typeof value === "number") {
                return value.toString();
            } else if (utils.isArray(value)) {
                if (value.length === 0) return "[]";
                
                // Simple case: short list of simple values
                if (value.length <= 3 && value.every(function(item) { 
                    return typeof item !== "object" && String(item).length < 15; 
                })) {
                    return "[" + value.join(", ") + "]";
                }
                
                // Complex case: return YAML notation
                var result = "\n";
                for (var i = 0; i < value.length; i++) {
                    result += indentStr + "- " + stringifyValue(value[i], indent + 2) + "\n";
                }
                return result.substring(0, result.length - 1); // Remove last newline
            } else if (typeof value === "string") {
                if (value.indexOf("\n") !== -1) {
                    // Multiline string
                    var result = "|\n";
                    var lines = value.split(/\r\n|\r|\n/);
                    for (var i = 0; i < lines.length; i++) {
                        result += indentStr + lines[i] + "\n";
                    }
                    return result.substring(0, result.length - 1);
                } else if (utils.needsQuotes(value)) {
                    return "\"" + utils.escapeString(value) + "\"";
                } else {
                    return value;
                }
            } else if (typeof value === "object") {
                // Nested object
                var result = "\n";
                for (var prop in value) {
                    if (value.hasOwnProperty(prop)) {
                        result += indentStr + prop + ": " + stringifyValue(value[prop], indent + 2) + "\n";
                    }
                }
                return result.substring(0, result.length - 1);
            }
            
            // Default case
            return String(value);
        }
        
        // Public API
        return {
            parse: parse,
            stringify: stringify,
            utils: utils
        };
    })();
    
    /**
     * @namespace PandocMapper
     * @description Maps Pandoc YAML metadata to BookCreator format
     */
    var PandocMapper = (function() {
        
        /**
         * Maps Pandoc metadata to BookCreator format
         * @param {Object} pandocData - Parsed Pandoc YAML data
         * @return {Object} BookCreator compatible metadata
         */
        function mapToBookCreator(pandocData) {
            var result = {};
            
            // Handle title - can be array or string
            if (pandocData.title) {
                if (isArray(pandocData.title)) {
                    // Complex title structure
                    for (var i = 0; i < pandocData.title.length; i++) {
                        var titleItem = pandocData.title[i];
                        if (titleItem && typeof titleItem === "object") {
                            if (titleItem.type === "main" && titleItem.text) {
                                result.title = titleItem.text;
                            } else if (titleItem.type === "subtitle" && titleItem.text) {
                                result.subtitle = titleItem.text;
                            }
                        }
                    }
                } else {
                    // Simple string title
                    result.title = pandocData.title;
                    // Ne PAS traiter subtitle ici si on a un title simple
                }
            }
            
            // Handle creator - can be array or string  
            if (pandocData.creator) {
                if (isArray(pandocData.creator)) {
                    // Complex creator structure
                    for (var j = 0; j < pandocData.creator.length; j++) {
                        var creatorItem = pandocData.creator[j];
                        if (creatorItem && typeof creatorItem === "object") {
                            if (creatorItem.role === "author" && creatorItem.text) {
                                result.author = creatorItem.text;
                            }
                        }
                    }
                } else {
                    // Simple string creator
                    result.author = pandocData.creator;
                }
            }
            
            // Handle identifier for ISBN
            if (pandocData.identifier && isArray(pandocData.identifier)) {
                for (var k = 0; k < pandocData.identifier.length; k++) {
                    var identifierItem = pandocData.identifier[k];
                    if (identifierItem && typeof identifierItem === "object") {
                        if (identifierItem.scheme === "ISBN" && identifierItem.text) {
                            if (!result.isbnEbook) {
                                result.isbnEbook = identifierItem.text;
                            }
                        }
                    }
                }
            }
            
            // Direct field mappings
            if (pandocData["isbn-print"]) result.isbnPrint = pandocData["isbn-print"];
            if (pandocData["isbn-ebook"]) result.isbnEbook = pandocData["isbn-ebook"];
            
            if (pandocData["published-print"]) {
                result.printDate = pandocData["published-print"];
            } else if (pandocData.date) {
                result.printDate = pandocData.date;
            }
            
            if (pandocData["translator-display"]) {
                result.translation = pandocData["translator-display"];
            }
            
            if (pandocData["critical-display"]) {
                result.critical = pandocData["critical-display"];
            }
            
            if (pandocData["cover-note"]) {
                result.coverCredit = pandocData["cover-note"];
            }
            
            // Other field mappings
            if (pandocData.editions) result.editions = pandocData.editions;
            if (pandocData.funding) result.funding = pandocData.funding;
            if (pandocData.rights) result.rights = pandocData.rights;
            if (pandocData.price) result.price = pandocData.price;
            if (pandocData.publisher) result.publisher = pandocData.publisher;
            if (pandocData.lang) result.language = pandocData.lang;
            if (pandocData["original-title"]) result.originalTitle = pandocData["original-title"];
            
            // SUPPRIMEZ ou CORRIGEZ ces lignes problématiques :
            // Ne les exécutez QUE si les champs n'ont pas déjà été définis
            // ET si les champs simples existent vraiment dans pandocData
            
            // Pour subtitle : seulement si pas déjà défini ET si le champ existe vraiment
            if (!result.subtitle && pandocData.subtitle && pandocData.subtitle !== undefined) {
                result.subtitle = pandocData.subtitle;
            }
            
            // Pour author : seulement si pas déjà défini ET si le champ existe vraiment  
            if (!result.author && pandocData.author && pandocData.author !== undefined) {
                result.author = pandocData.author;
            }
            
            return result;
        }
        
        // Public API
        return {
            mapToBookCreator: mapToBookCreator
        };
    })();
    
    /**
     * @namespace I18n
     * @description Internationalization module for UI translations
     */
    var I18n = (function() {
        // Current language
        var currentLanguage = detectInDesignLanguage();
        
        // Translation dictionaries
        var translations = {
            'en': {
                // Main interface
                'title': 'BookCreator v%s',
                'fileNamePrefix': 'File name prefix:',
                'chapterCount': 'Number of chapters:',
                'templates': 'Templates',
                'chapterTemplate': 'Chapter Template',
                'beforeTemplate': 'Template Before Chapters',
                'afterTemplate': 'Template After Chapters',
                'coverTemplate': 'Cover Template',
                'noFileSelected': 'No file selected.',
                'filesSelected': '%d files selected.',
                'bookInformation': 'Book Information',
                'injectMarkdown': 'Inject Markdown',
                'cancel': 'Cancel',
                'createBook': 'Create Book',
                'chooseDestionationFolder': 'Choose destination folder',
                'bookSuffix': 'Book',
                'searchLocations': 'Searched locations:%s',
                
                // Book information window
                'bookAuthor': 'Book Author:',
                'bookTitle': 'Book Title:',
                'subtitle': 'Subtitle:',
                'isbnPrint': 'ISBN Print:',
                'isbnEbook': 'ISBN Ebook:',
                'price': 'Price:',
                'originalTitle': 'Original Title:',
                'prefix': 'Prefix',
                'printDate': 'Print Date:',
                'critical': 'Critical Apparatus:',
                'translation': 'Translation:',
                'coverCredit': 'Cover Credit:',
                'editions': 'Editions:',
                'funding': 'Funding:',
                'rights': 'Rights / License:',
                'importYAML': 'Import YAML',
                'exportYAML': 'Export YAML',
                'save': 'Save',
                'markdownDetection': 'Markdown Detection:',
                'inputFilesDetected': 'input-files detected',
                'inputFilesNotDetected': 'input-files not detected',
                'noYAMLLoaded': 'No YAML file loaded',
                
                // Success and error messages
                'bookGenerated': 'Book successfully generated!',
                'yamlImportCompleted': 'YAML import completed.',
                'yamlExportCompleted': 'YAML export completed.',
                'fileExists': 'File %s already exists. Do you want to replace it?',
                'destinationFolderNotExist': 'Destination folder does not exist.',
                'markdownFileNotFound': 'Markdown file not found: %s',
                'noTextFrameFound': 'No text frame found in document.',
                'error': 'Error',
                
                // Default prefixes
                'originalTitlePrefix': 'Original Title: ',
                'coverCreditPrefix': 'Cover: ',
                
                // Validation
                'bookNameRequired': 'Book name is required.',
                'chapterCountPositive': 'Chapter count must be a positive number.',
                'chapterTemplateRequired': 'A chapter template is required.',
                'invalidPrintISBN': 'Invalid Print ISBN: %s',
                'invalidEbookISBN': 'Invalid Ebook ISBN: %s',
                
                // Language selector
                'languageChanged': 'Language has been changed. Please restart the script to apply changes.'
            },
            'fr': {
                // Main interface
                'title': 'BookCreator v%s',
                'fileNamePrefix': 'Pr\u00E9fixe de nom de fichier :',
                'chapterCount': 'Nombre de chapitres :',
                'templates': 'Mod\u00E8les',
                'chapterTemplate': 'Mod\u00E8le de chapitre',
                'beforeTemplate': 'Mod\u00E8le avant chapitres',
                'afterTemplate': 'Mod\u00E8le apr\u00E8s chapitres',
                'coverTemplate': 'Mod\u00E8le de couverture',
                'noFileSelected': 'Aucun fichier s\u00E9lectionn\u00E9.',
                'filesSelected': '%d fichiers s\u00E9lectionn\u00E9s.',
                'bookInformation': 'Informations du livre',
                'injectMarkdown': 'Injecter Markdown',
                'cancel': 'Annuler',
                'createBook': 'Cr\u00E9er le livre',
                'chooseDestionationFolder': 'Choisir le dossier de destination',
                'bookSuffix': 'Livre',
                'searchLocations': 'Emplacements recherchés:%s',
                
                // Book information window
                'bookAuthor': 'Auteur du livre :',
                'bookTitle': 'Titre du livre :',
                'subtitle': 'Sous-titre :',
                'isbnPrint': 'ISBN Imprim\u00E9 :',
                'isbnEbook': 'ISBN Ebook :',
                'price': 'Prix :',
                'originalTitle': 'Titre original :',
                'prefix': 'Pr\u00E9fixe',
                'printDate': 'Date d\'impression :',
                'critical': 'Appareil critique :',
                'translation': 'Traduction :',
                'coverCredit': 'Cr\u00E9dit couverture :',
                'editions': '\u00C9ditions :',
                'funding': 'Financement :',
                'rights': 'Droits / Licence :',
                'importYAML': 'Importer YAML',
                'exportYAML': 'Exporter YAML',
                'save': 'Enregistrer',
                'markdownDetection': 'D\u00E9tection Markdown :',
                'inputFilesDetected': 'fichiers d\'entr\u00E9e d\u00E9tect\u00E9s',
                'inputFilesNotDetected': 'fichiers d\'entr\u00E9e non d\u00E9tect\u00E9s',
                'noYAMLLoaded': 'Aucun fichier YAML charg\u00E9',
                
                // Success and error messages
                'bookGenerated': 'Livre g\u00E9n\u00E9r\u00E9 avec succ\u00E8s !',
                'yamlImportCompleted': 'Import YAML termin\u00E9.',
                'yamlExportCompleted': 'Export YAML termin\u00E9.',
                'fileExists': 'Le fichier %s existe d\u00E9j\u00E0. Voulez-vous le remplacer ?',
                'destinationFolderNotExist': 'Le dossier de destination n\'existe pas.',
                'markdownFileNotFound': 'Fichier Markdown non trouv\u00E9 : %s',
                'noTextFrameFound': 'Aucun cadre de texte trouv\u00E9 dans le document.',
                'error': 'Erreur',
                
                // Default prefixes
                'originalTitlePrefix': 'Titre original\u2009: ',
                'coverCreditPrefix': 'Couverture\u2009: ',
                
                // Validation
                'bookNameRequired': 'Le nom du livre est requis.',
                'chapterCountPositive': 'Le nombre de chapitres doit \u00EAtre un nombre positif.',
                'chapterTemplateRequired': 'Un mod\u00E8le de chapitre est requis.',
                'invalidPrintISBN': 'ISBN Imprim\u00E9 invalide : %s',
                'invalidEbookISBN': 'ISBN Ebook invalide : %s',
                
                // Language selector
                'languageChanged': 'La langue a \u00E9t\u00E9 chang\u00E9e. Veuillez red\u00E9marrer le script pour appliquer les modifications.'
            }
        };
        
        /**
         * Gets a translated string with optional substitutions
         * @param {string} key - Translation key
         * @param {...*} args - Arguments for substitutions
         * @return {string} Translated string
         */
        function __(key) {
            var lang = currentLanguage;
            var langDict = translations[lang] || translations['en'];
            var str = langDict[key] || translations['en'][key] || key;
            
            // If additional arguments are provided, use them for formatting
            if (arguments.length > 1) {
                var args = Array.prototype.slice.call(arguments, 1);
                str = str.replace(/%[sdx]/g, function(match) {
                    if (!args.length) return match;
                    var arg = args.shift();
                    switch (match) {
                        case '%s': return String(arg);
                        case '%d': return parseInt(arg, 10);
                        case '%x': return '0x' + parseInt(arg, 10).toString(16);
                        default: return match;
                    }
                });
            }
            
            return str;
        }
        
        /**
         * Changes the current language
         * @param {string} lang - Language code ('fr' or 'en')
         */
        function setLanguage(lang) {
            if (translations[lang]) {
                currentLanguage = lang;
            }
        }
        
        /**
         * Gets the current language
         * @return {string} Current language code
         */
        function getLanguage() {
            return currentLanguage;
        }
        
        /**
         * Detects the language of the InDesign interface
         * @return {string} Language code ('fr' or 'en')
         */
        function detectInDesignLanguage() {
            try {
                // Debug info to trace execution
                $.writeln("Attempting to detect InDesign language...");
                
                // Get localization string using the full app object
                var locale = "";
                
                // Try different methods to access locale
                if (typeof app !== 'undefined' && app.hasOwnProperty('locale')) {
                    locale = String(app.locale);
                    $.writeln("Detected locale: " + locale);
                } else if (typeof app !== 'undefined' && app.hasOwnProperty('languageAndRegion')) {
                    locale = String(app.languageAndRegion);
                    $.writeln("Detected languageAndRegion: " + locale);
                } else {
                    $.writeln("Could not access InDesign locale properties");
                    return 'en'; // Default to English
                }
                
                // Convert locale to lowercase for case-insensitive comparison
                locale = locale.toLowerCase();
                
                // Debug the detected locale
                $.writeln("Normalized locale: " + locale);
                
                // Check for French locales
                if (locale.indexOf('fr') !== -1) {
                    $.writeln("French locale detected, setting language to fr");
                    return 'fr';
                } else {
                    // Default to English for any other locale
                    $.writeln("Non-French locale detected, setting language to en");
                    return 'en';
                }
            } catch (e) {
                // Log detailed error information
                $.writeln("Error detecting language: " + e);
                $.writeln("Error details: " + e.message);
                if (e.line) $.writeln("Error line: " + e.line);
                if (e.stack) $.writeln("Error stack: " + e.stack);
                
                // In case of error, use English by default
                return 'en';
            }
        }
        
        // Then set current language
        var currentLanguage = detectInDesignLanguage();
        
        // Public API
        return {
            __: __,
            setLanguage: setLanguage,
            getLanguage: getLanguage,
            detectLanguage: detectInDesignLanguage
        };
    })();
    
    /**
     * @namespace TextUtils
     * @description Text formatting utilities with simple markdown formatting
     */
    var TextUtils = {
        /**
         * Applies formatted text to a text frame with basic markdown formatting
         * @param {TextFrame} textFrame - InDesign text frame to apply text to
         * @param {string} text - Raw text content with markdown
         * @param {Document} doc - Parent InDesign document
         */
        applyFormattedText: function(textFrame, text, doc) {
            if (!text) {
                textFrame.contents = "";
                return;
            }
            
            // Process basic replacements
            var processedText = text.replace(/<br\s*\/?>/gi, "\n");
            processedText = processedText.replace(/[ ]{2,}$/mg, "\n");
            
            // Apply text to frame
            textFrame.contents = processedText;
            
            // Apply markdown formatting if italic style exists
            this.formatItalicMarkdown(textFrame, doc);
        },
        
        /**
         * Formats text between ** markers as italic if italic character style exists
         * @param {TextFrame} textFrame - Target text frame
         * @param {Document} doc - InDesign document
         */
        formatItalicMarkdown: function(textFrame, doc) {
            try {
                // Check if italic character style exists
                var italicStyle = doc.characterStyles.itemByName("Italic");
                if (!italicStyle.isValid) {
                    // No italic style found, leave text as-is with * markers
                    return;
                }
                
                // Clear find/change preferences
                app.findGrepPreferences = app.changeGrepPreferences = null;
                
                // Find text between * markers and replace directly with GREP
                app.findGrepPreferences.findWhat = "\\*(.+?)\\*";
                app.changeGrepPreferences.changeTo = "$1"; // $1 = premier groupe capturé (sans les *)
                app.changeGrepPreferences.appliedCharacterStyle = italicStyle;
                
                // Apply changes only in this text frame
                textFrame.changeGrep();
                
                // Clear preferences
                app.findGrepPreferences = app.changeGrepPreferences = null;
                
            } catch (e) {
                $.writeln("Warning: Error applying italic formatting: " + e.message);
            }
        }
    };
    
    /**
     * @namespace PageOverflow
     * @description Handles text overflow by adding pages automatically
     */
    var PageOverflow = {
        /**
         * Processes overflowing text frames by adding pages and linked frames
         * @param {Document} doc - InDesign document to process
         * @return {boolean} True if pages were added, false otherwise
         */
        processOverflow: function(doc) {
            if (!doc || !doc.isValid) return false;
            
            var textFrame = null;
            var lastPage = doc.pages[-1]; // Get last page
            
            // Check the last page first for overflowing frames
            for (var i = 0; i < lastPage.textFrames.length; i++) {
                if (lastPage.textFrames[i].overflows) {
                    textFrame = lastPage.textFrames[i];
                    break;
                }
            }
            
            // If no overflowing frame found on last page, check all pages
            if (textFrame === null) {
                for (var j = 0; j < doc.pages.length; j++) {
                    var currentPage = doc.pages[j];
                    for (var k = 0; k < currentPage.textFrames.length; k++) {
                        if (currentPage.textFrames[k].overflows) {
                            textFrame = currentPage.textFrames[k];
                            break;
                        }
                    }
                    if (textFrame !== null) break;
                }
            }
            
            // If no overflowing frame found in entire document
            if (textFrame === null) {
                return false; // Nothing to process
            }
            
            var pageCount = 0; // Counter to avoid infinite loops
            var maxPages = 500; // Safety limit
    
            while (textFrame.overflows && pageCount < maxPages) {
                // Create a new page at the end
                var newPage = doc.pages.add(LocationOptions.AFTER, doc.pages[-1]);
                pageCount++;
                
                // Determine if it's a left or right page
                var isLeftPage = newPage.side == PageSideOptions.LEFT_HAND;
                
                // Set margins based on page type
                var topMargin = newPage.marginPreferences.top;
                var bottomMargin = doc.documentPreferences.pageHeight - newPage.marginPreferences.bottom;
                
                var leftMargin, rightMargin;
                
                if (doc.documentPreferences.facingPages) {
                    // For facing pages layouts
                    if (isLeftPage) {
                        leftMargin = newPage.marginPreferences.right; // Outside margin for left page
                        rightMargin = doc.documentPreferences.pageWidth - newPage.marginPreferences.left; // Inside margin for left page
                    } else {
                        leftMargin = newPage.marginPreferences.left; // Inside margin for right page
                        rightMargin = doc.documentPreferences.pageWidth - newPage.marginPreferences.right; // Outside margin for right page
                        
                        // For right pages, apply offset of page width for correct positioning
                        leftMargin += doc.documentPreferences.pageWidth;
                        rightMargin += doc.documentPreferences.pageWidth;
                    }
                } else {
                    // For non-facing pages
                    leftMargin = newPage.marginPreferences.left;
                    rightMargin = doc.documentPreferences.pageWidth - newPage.marginPreferences.right;
                }
    
                // Create a new text frame on the new page with appropriate margins
                var newFrame = newPage.textFrames.add({
                    geometricBounds: [topMargin, leftMargin, bottomMargin, rightMargin]
                });
    
                // Link the previous text frame to the new one
                textFrame.nextTextFrame = newFrame;
    
                // Continue with the new frame
                textFrame = newFrame;
            }
    
            // If document has facing pages and ends with a right page, add an empty left page
            if (doc.documentPreferences.facingPages && !textFrame.overflows) {
                var lastPage = doc.pages[-1];
                if (lastPage.side == PageSideOptions.RIGHT_HAND) {
                    doc.pages.add(LocationOptions.AFTER, lastPage);
                }
            }
            
            return pageCount > 0; // Return true if pages were added
        }
    };
    
    /**
     * Book class for managing InDesign book creation
     * @class
     * @param {string} name - Book name
     * @param {Object} info - Book metadata
     */
    function Book(name, info) {
        this.name = name;
        this.info = info || {};
        this.templates = {
            before: [],
            frontmatter: null,
            bodymatter: null,
            backmatter: null,
            after: [],
            cover: null
        };
        this.templateFolder = null;
        this.chapterCount = 0;
        
        // Display options (checked by default)
        this.displayOptions = {
            showOriginalTitleLabel: true,
            originalTitleLabelText: I18n.__('originalTitlePrefix'), 
            showCoverCreditLabel: true,
            coverCreditLabelText: I18n.__('coverCreditPrefix')
        };
        
        // Markdown options
        this.markdownOptions = {
            injectMarkdown: false,
            yamlPath: null,
            yamlMeta: null,
            hasInputFiles: false
        };
        
        /**
         * Validates book data before generation
         * @return {Object} Validation result with valid status and message
         */
        this.validate = function() {
            if (!this.name) {
                return { valid: false, message: I18n.__('bookNameRequired') };
            }
            
            // Validation basée sur l'auto-découverte des fichiers Markdown
            if (this.markdownOptions.injectMarkdown && this.markdownOptions.yamlPath) {
                var autoDiscoveredFiles = this._getAllMarkdownFilesFromFolder();
                if (autoDiscoveredFiles.length === 0) {
                    return { valid: false, message: "No Markdown files found in text folder." };
                }
                // Auto-définir le nombre de chapitres basé sur les fichiers MD découverts
                this.chapterCount = autoDiscoveredFiles.length;
            } else {
                // Mode manuel si pas de Markdown
                if (!this.chapterCount || isNaN(parseInt(this.chapterCount)) || parseInt(this.chapterCount) <= 0) {
                    return { valid: false, message: I18n.__('chapterCountPositive') };
                }
            }
            
            if (!this.templateFolder) {
                return { valid: false, message: "A templates folder is required." };
            }
            
            if (!this.templates.bodymatter) {
                return { valid: false, message: "No Bodymatter template found in folder." };
            }
            
            // Validate ISBNs if present
            if (this.info.isbnPrint) {
                var printResult = BookUtils.ISBN.validate(this.info.isbnPrint);
                if (!printResult.valid) {
                    return { valid: false, message: I18n.__('invalidPrintISBN', printResult.message) };
                }
            }
            
            if (this.info.isbnEbook) {
                var ebookResult = BookUtils.ISBN.validate(this.info.isbnEbook);
                if (!ebookResult.valid) {
                    return { valid: false, message: I18n.__('invalidEbookISBN', ebookResult.message) };
                }
            }
            
            return { valid: true };
        };
        
        /**
         * Generates the InDesign book and related documents
         * @param {Folder} folder - Destination folder
         * @return {Book|boolean} InDesign book object or false if generation failed
         */
        this.generate = function(folder, mainWindow) {
            if (!folder || !folder.exists) {
                throw new Error(I18n.__('destinationFolderNotExist'));
            }
            
            var result = this.validate();
            if (!result.valid) {
                throw new Error(result.message);
            }
            
            // Calculate total steps based on discovered templates and MD files
            var allMdFiles = this._getAllMarkdownFilesFromFolder();
            var totalSteps = this.templates.before.length + 
                             allMdFiles.length + 
                             this.templates.after.length + 
                             (this.templates.cover ? 1 : 0) + 
                             2;
            
            try {
                // Fermer la fenêtre principale avant la génération
                if (mainWindow) {
                    mainWindow.close();
                }
                
                // Créer la barre de progression APRÈS avoir fermé la fenêtre
                var progress = new ProgressBar(
                    I18n.getLanguage() === 'fr' ? "G\u00E9n\u00E9ration du livre..." : "Generating book...",
                    totalSteps
                );
                
                var currentStep = 0;
                
                // Créer le fichier InDesign book
                progress.update(currentStep++, 
                    I18n.getLanguage() === 'fr' ? "Cr\u00E9ation du fichier livre..." : "Creating book file...");
                
                var bookFileName = this.name.toUpperCase() + '-' + I18n.__('bookSuffix') + '.indb';
                var bookFile = new File(folder.fsName + '/' + bookFileName);
                
                if (bookFile.exists) {
                    progress.close(); // Fermer temporairement pour le dialogue
                    var overwrite = confirm(I18n.__('fileExists', bookFileName));
                    if (!overwrite) {
                        return false;
                    }
                    progress.show(); // Réafficher après le dialogue
                }
                
                progress.show(); // Montrer la barre maintenant
                
                var book = app.books.add(bookFile);
                // Petite pause pour s'assurer que le livre est bien initialisé
                app.doScript(function(){}, ScriptLanguage.JAVASCRIPT, [], UndoModes.FAST_ENTIRE_SCRIPT);
                var prefix = this.name.toUpperCase() + '-';
                
                // Generate documents before chapters
                for (var i = 0; i < this.templates.before.length; i++) {
                    progress.update(
                        currentStep++, 
                        I18n.getLanguage() === 'fr' ? "Traitement des modèles avant chapitres..." : "Processing templates before chapters...",
                        this.templates.before[i].name
                    );
                    
                    this._generateDocument(
                        folder, 
                        this.templates.before[i], 
                        this.templates.before[i].name.replace(/^.*?-/, prefix), 
                        true, 
                        book
                    );
                }
                
                // Generate all documents based on Markdown files and appropriate templates
                var allMdFiles = this._getAllMarkdownFilesFromFolder();
                var currentStep = this.templates.before.length; // Ajuster le compteur
                
                // Generate documents before chapters
                for (var b = 0; b < this.templates.before.length; b++) {
                    progress.update(
                        currentStep++, 
                        I18n.getLanguage() === 'fr' ? "Traitement des templates avant..." : "Processing before templates...",
                        this.templates.before[b].name
                    );
                    
                    this._generateDocument(
                        folder, 
                        this.templates.before[b], 
                        this.templates.before[b].name.replace(/^.*?-/, prefix), 
                        true, 
                        book
                    );
                }
                
                // Generate all Markdown-based documents
                for (var m = 0; m < allMdFiles.length; m++) {
                    var currentMdFile = allMdFiles[m];
                    var descriptiveName = currentMdFile.replace(/^\d+-/, "").replace(/\.md$/, "");
                    var templateType = this._determineTemplateType(currentMdFile);
                    
                    // Select appropriate template with priority matching
                    var selectedTemplate = this._selectBestTemplate(currentMdFile, templateType);
                    
                    if (selectedTemplate) {
                        progress.update(
                            currentStep++, 
                            I18n.getLanguage() === 'fr' ? 
                                "Cr\u00E9ation de " + descriptiveName + " (" + (m + 1) + " sur " + allMdFiles.length + ")..." :
                                "Creating " + descriptiveName + " (" + (m + 1) + " of " + allMdFiles.length + ")...",
                            descriptiveName + ".indd"
                        );
                        
                        var newName = selectedTemplate.name.replace(/^.*?-/, prefix).replace('.indd', '-' + descriptiveName + '.indd');
                        
                        this._generateDocument(
                            folder,
                            selectedTemplate,
                            newName,
                            true,
                            book,
                            m  // Index pour le matching MD
                        );
                    }
                }
                
                // Generate documents after chapters
                for (var a = 0; a < this.templates.after.length; a++) {
                    progress.update(
                        currentStep++, 
                        I18n.getLanguage() === 'fr' ? "Traitement des templates apr\u00E8s..." : "Processing after templates...",
                        this.templates.after[a].name
                    );
                    
                    this._generateDocument(
                        folder,
                        this.templates.after[a],
                        this.templates.after[a].name.replace(/^.*?-/, prefix),
                        true,
                        book
                    );
                }
                
                // Generate cover (separate from book)
                if (this.templates.cover) {
                    progress.update(
                        currentStep++, 
                        I18n.getLanguage() === 'fr' ? "Cr\u00E9ation de la couverture..." : "Creating cover...",
                        "Cover.indd"
                    );
                    
                    this._generateDocument(
                        folder,
                        this.templates.cover,
                        prefix + 'Cover.indd',
                        false,  // Important : false pour ne pas l'inclure dans le livre
                        book
                    );
                }
                
                // Save book
                progress.update(currentStep++, 
                    I18n.getLanguage() === 'fr' ? "Sauvegarde du livre..." : "Saving book...");
                book.save();
                
                progress.close();
                return book;
                
            } catch (e) {
                if (typeof progress !== 'undefined') {
                    progress.close();
                }
                LogManager.logError("Error generating book", e);
                return false;
            }
        };
        
        /**
         * Generates a single document from template
         * @param {Folder} folder - Destination folder
         * @param {File} template - Template file
         * @param {string} newName - New document name
         * @param {boolean} includeInBook - Whether to include in book
         * @param {Book} book - InDesign book object
         * @return {boolean} Success status
         * @private
         */
        this._generateDocument = function(folder, template, newName, includeInBook, book, chapterIndex) {
            var destFile = new File(folder.fsName + '/' + newName);
            var doc = null;
            
            try {
                // Copy template
                if (!template.copy(destFile.fsName)) {
                    throw new Error("Failed to copy template file");
                }
                
                // Open document
                doc = app.open(destFile, false);
                
                // Add custom variables
                for (var key in this.info) {
                    if (this.info.hasOwnProperty(key)) {
                        try {
                            BookUtils.Document.addCustomVariable(doc, key, this.info[key]);
                        } catch (e) {
                            // Continue if variable creation fails
                            $.writeln("Warning: Could not add variable " + key + ": " + e.message);
                        }
                    }
                }
                
                // Replace text placeholders
                try {
                    BookUtils.Document.replaceTextPlaceholders(doc, this.info, this.displayOptions);
                } catch (e) {
                    $.writeln("Warning: Error replacing text placeholders: " + e.message);
                }
                
                // Replace EAN13 placeholders
                try {
                    BookUtils.Document.replaceEAN13Placeholders(doc, this.info.isbnPrint, this.info.isbnEbook);
                } catch (e) {
                    $.writeln("Warning: Error replacing EAN13 placeholders: " + e.message);
                }
                
                // Inject Markdown if needed
                if (this.markdownOptions.injectMarkdown && this.markdownOptions.hasInputFiles) {
                    try {
                        this._injectMarkdownContent(doc, chapterIndex);
                        
                        if (doc.isValid) {
                            PageOverflow.processOverflow(doc);
                        }
                    } catch (e) {
                        $.writeln("Warning: Error processing Markdown content: " + e.message);
                    }
                }
                
                // Save document
                doc.save(destFile);
                doc.close();
                doc = null; // Clear reference
                
                // Add to book if needed - with retry mechanism
                if (includeInBook && book) {
                    var maxRetries = 3;
                    var retryCount = 0;
                    var addedToBook = false;
                    
                    while (retryCount < maxRetries && !addedToBook) {
                        try {
                            // Small delay before adding to book
                            app.doScript(function(){}, ScriptLanguage.JAVASCRIPT, [], UndoModes.FAST_ENTIRE_SCRIPT);
                            
                            book.bookContents.add(destFile);
                            addedToBook = true;
                        } catch (e) {
                            retryCount++;
                            $.writeln("Retry " + retryCount + " adding to book: " + e.message);
                            
                            if (retryCount >= maxRetries) {
                                // Log final error but don't fail the generation
                                $.writeln("Warning: Could not add " + newName + " to book after " + maxRetries + " attempts: " + e.message);
                                break;
                            }
                        }
                    }
                }
                
                return true;
                
            } catch (e) {
                // Clean up document if still open
                if (doc && doc.isValid) {
                    try {
                        doc.close(SaveOptions.NO);
                    } catch (closeError) {
                        // Ignore close errors
                    }
                }
                
                LogManager.logError("Error generating document " + newName, e);
                return false;
            }
        };
        
        /**
         * Injects Markdown content into a document while preserving line breaks
         * @param {Document} doc - InDesign document
         * @return {boolean} Success status
         * @private
         */
        this._injectMarkdownContent = function(doc, chapterIndex) {
            try {
                // Check parameters
                if (!doc) {
                    LogManager.logError("ERROR: InDesign document not defined!");
                    return false;
                }
                
                if (!this.markdownOptions || !this.markdownOptions.yamlPath) {
                    LogManager.logError("ERROR: YAML path not defined!");
                    return false;
                }
                
                // Check that YAML file exists
                var yamlFile = File(this.markdownOptions.yamlPath);
                if (!yamlFile.exists) {
                    LogManager.logError("ERROR: YAML file does not exist:\n" + this.markdownOptions.yamlPath);
                    return false;
                }
                
                // Direct injection implementation, without calling external script
                
                // 1. Read YAML
                function localTrim(str) {
                    return str.replace(/^\s+|\s+$/g, '');
                }
        
                function parseYaml(yamlString) {
                    if (!yamlString) return {};
                
                    // Retire un éventuel front-matter délimité par --- ... ---
                    // (gère \n et \r\n)
                    var s = String(yamlString);
                    var m = s.match(/^\s*---\s*\r?\n([\s\S]*?)\r?\n---\s*$/);
                    var body = m ? m[1] : s;
                
                    try {
                        return YAMLParser.parse(body);
                    } catch (e) {
                        // Sécurise en cas d’erreur de parsing pour ne pas bloquer l’exécution
                        try {
                            if (typeof LogManager !== "undefined" && LogManager.logError) {
                                LogManager.logError("YAML parse error in _injectMarkdownContent: " + e);
                            }
                        } catch (_e) {}
                        return {};
                    }
                }
        
                // Read YAML file content
                yamlFile.open("r");
                var yamlContent = yamlFile.read();
                yamlFile.close();
        
                var yamlData = parseYaml(yamlContent);
                var inputFiles = this._getAllMarkdownFilesFromFolder();
                
                // Base directory of YAML file
                var configDir = yamlFile.parent;
                
                // Determine project root directory
                var projectDir;
                var configDirName = configDir.name.toLowerCase();
                
                // Check if YAML is in a subdirectory or at root
                if (configDirName === "config" || configDirName === "configs" || configDirName === "metadata") {
                    // YAML is in a config subdirectory, project is parent
                    projectDir = configDir.parent;
                } else {
                    // YAML is likely at project root
                    projectDir = configDir;
                }
                
                // Define possible text folder names with different conventions
                var textFolderVariants = ["text", "Text", "texts", "Texts", "texte", "Texte", "textes", "Textes", "md", "MD", "markdown", "Markdown"];
                
                // Create an array to store all searchable text folders
                var textFolders = [];
                
                // Check for each variant if the folder exists
                for (var v = 0; v < textFolderVariants.length; v++) {
                    var folderVariant = new Folder(projectDir.fsName + "/" + textFolderVariants[v]);
                    if (folderVariant.exists) {
                        textFolders.push(folderVariant);
                    }
                }
        
                if (!inputFiles || inputFiles.length === 0) {
                    alert("No Markdown files found in text folder for auto-injection.");
                    return false;
                }
        
                // 2. Find matching Markdown file
                var docName = doc.name;
                // Utiliser l'index du chapitre pour le matching (si fourni)
                var mdFileName;
                if (typeof chapterIndex !== "undefined" && chapterIndex !== null) {
                    mdFileName = this._findMatchingMarkdownFile(chapterIndex, inputFiles);
                } else {
                    // Fallback : essayer de déduire l'index depuis le nom du document
                    var chapterMatch = doc.name.match(/_(\d+)\.indd$/);
                    if (chapterMatch) {
                        var deducedIndex = parseInt(chapterMatch[1], 10) - 1;
                        mdFileName = this._findMatchingMarkdownFile(deducedIndex, inputFiles);
                    } else {
                        return false;
                    }
                }
                if (!mdFileName) {
                    // Return silently without error
                    return false;
                }
        
                // Search for the file in possible locations (by priority order)
                var mdFile = null;
                var searchLocations = [];
                
                // 1. First look in all text folder variants
                for (var t = 0; t < textFolders.length; t++) {
                    var textFolderPath = textFolders[t].fsName + "/" + mdFileName;
                    var textFolderFile = File(textFolderPath);
                    searchLocations.push(textFolderPath);
                    
                    if (textFolderFile.exists) {
                        mdFile = textFolderFile;
                        break; // Exit loop if file is found
                    }
                }
                
                // If not found in text folders, continue with other locations
                if (!mdFile) {
                    // 2. Look in the same folder as the YAML (previous behavior)
                    var configDirPath = configDir.fsName + "/" + mdFileName;
                    var configDirFile = File(configDirPath);
                    searchLocations.push(configDirPath);
                    
                    if (configDirFile.exists) {
                        mdFile = configDirFile;
                    } else {
                        // 3. Look in the project root folder
                        var projectDirPath = projectDir.fsName + "/" + mdFileName;
                        var projectDirFile = File(projectDirPath);
                        searchLocations.push(projectDirPath);
                        
                        if (projectDirFile.exists) {
                            mdFile = projectDirFile;
                        }
                    }
                }
                
                // If not found, display a multilingual error
                if (!mdFile) {
                    var searchPathsMsg = "";
                    for (var i = 0; i < searchLocations.length; i++) {
                        searchPathsMsg += "\n- " + searchLocations[i];
                    }
                    
                    // Use I18n to display a localized error message
                    alert(I18n.__('markdownFileNotFound', mdFileName) + 
                          "\n\n" + I18n.__('searchLocations', searchPathsMsg));
                    return false;
                }
        
                // IMPORTANT: Read Markdown file in binary mode to preserve line breaks exactly
                mdFile.encoding = "UTF-8"; // First set encoding for proper character handling
                mdFile.open("r");
                var mdContent = mdFile.read();
                mdFile.close();
        
                // 3. Find target text frame
                var targetFrame = this._findTargetTextFrame(doc);
                if (!targetFrame) {
                    alert(I18n.__('noTextFrameFound'));
                    return false;
                }
        
                // 4. Inject content with preserved line breaks
                // Remove epub:type attributes (with optional spaces) and convert line breaks
                var cleanedContent = mdContent.replace(/\s*\{epub\s*:\s*type\s*=\s*[^}]+\}/gi, "");
                // CRITICAL: Use \r instead of \n for InDesign line breaks
                targetFrame.contents = cleanedContent.replace(/\n/g, "\r");
                
                // 5. Extract first H1 title and replace <<Document_Title>>
                try {
                    // Find first H1 title in text frame
                    app.findGrepPreferences = app.changeGrepPreferences = null;
                    app.findGrepPreferences.findWhat = "^#\\s+(.+?)$";
                    var h1Finds = doc.findGrep();
                    
                    if (h1Finds.length > 0) {
                        // Extract title text without initial #
                        var titleText = h1Finds[0].contents.replace(/^#\s+/, "");
                        
                        // Trim spaces manually
                        titleText = titleText.replace(/^\s+|\s+$/g, "");
                        
                        // Clean up the title
                        titleText = titleText.replace(/\\/g, ""); // Remove backslashes
                        titleText = titleText.replace(/\s*\[\^[\w\d]+\]\s*/g, ""); // Remove footnote references
                        titleText = titleText.replace(/[*_~]/g, ""); // Remove formatting markers (bold, italic, strikethrough)
                        titleText = titleText.replace(/\s{2,}/g, " "); // Replace multiple spaces with a single space
                        
                        // Replace <<Document_Title>> throughout document
                        app.findTextPreferences = app.changeTextPreferences = null;
                        app.findTextPreferences.findWhat = "<<Document_Title>>";
                        app.changeTextPreferences.changeTo = titleText;
                        doc.changeText();
                    }
                } catch(e) {
                    // Ignore errors to avoid blocking main process
                }
                
                return true;
                
            } catch (e) {
                alert("Error injecting Markdown: " + e.message + "\n" + e.line);
                return false;
            }
        };
        
        /**
         * Finds matching Markdown file based on chapter index (Pandoc approach)
         * @param {number} chapterIndex - Index of current chapter (0-based)
         * @param {Array} mdFiles - List of Markdown filenames
         * @return {string|null} Matching Markdown filename or null
         * @private
         */
        this._findMatchingMarkdownFile = function(chapterIndex, mdFiles) {
            if (!mdFiles || !isArray(mdFiles) || mdFiles.length === 0) {
                return null;
            }
            
            // Filtrer les fichiers MD valides
            var validFiles = [];
            for (var i = 0; i < mdFiles.length; i++) {
                var file = mdFiles[i];
                // Accepter les chaînes ET les nombres, convertir en chaîne  
                if (file && (typeof file === "string" || typeof file === "number") && String(file) !== "") {
                    validFiles.push(String(file));
                }
            }
            
            // Tri alphabétique (comme Pandoc)
            validFiles.sort();
            
            // Retourner le fichier correspondant à l'index
            if (chapterIndex >= 0 && chapterIndex < validFiles.length) {
                return validFiles[chapterIndex];
            }
            
            return null;
        };
        
        /**
         * Finds target text frame for content injection
         * @param {Document} doc - InDesign document
         * @return {TextFrame|null} Target text frame or null
         * @private
         */
        this._findTargetTextFrame = function(doc) {
            // First, look for frame with "content" label
            var pages = doc.pages;
            for (var i = 0; i < pages.length; i++) {
                var frames = pages[i].textFrames;
                for (var j = 0; j < frames.length; j++) {
                    try {
                        if (frames[j].label === "content" || frames[j].label === "contenu") {
                            return frames[j];
                        }
                    } catch(e) {}
                }
            }
            
            // If not found, use first text frame
            if (doc.pages.length > 0 && doc.pages[0].textFrames.length > 0) {
                return doc.pages[0].textFrames[0];
            }
            
            return null;
        };
        
        /**
         * Auto-discovers all Markdown files in the text folder
         * @return {Array} Sorted list of Markdown filenames
         * @private
         */
        this._getAllMarkdownFilesFromFolder = function() {
            if (!this.markdownOptions || !this.markdownOptions.yamlPath) {
                return [];
            }
            
            try {
                var yamlFile = File(this.markdownOptions.yamlPath);
                if (!yamlFile.exists) return [];
                
                var configDir = yamlFile.parent;
                var projectDir;
                var configDirName = configDir.name.toLowerCase();
                
                // Determine project root directory
                if (configDirName === "config" || configDirName === "configs" || configDirName === "metadata") {
                    projectDir = configDir.parent;
                } else {
                    projectDir = configDir;
                }
                
                var textFolderVariants = ["text", "Text", "texts", "Texts", "texte", "Texte", "textes", "Textes", "md", "MD", "markdown", "Markdown"];
                var allFiles = [];
                
                // Search in all possible text folder variants
                for (var v = 0; v < textFolderVariants.length; v++) {
                    var folderVariant = new Folder(projectDir.fsName + "/" + textFolderVariants[v]);
                    if (folderVariant.exists) {
                        var files = folderVariant.getFiles("*.md");
                        for (var f = 0; f < files.length; f++) {
                            if (files[f] instanceof File) {
                                allFiles.push(files[f].name);
                            }
                        }
                        break; // Use first text folder found
                    }
                }
                
                // If no specialized folder found, search in project root
                if (allFiles.length === 0) {
                    var rootFiles = projectDir.getFiles("*.md");
                    for (var r = 0; r < rootFiles.length; r++) {
                        if (rootFiles[r] instanceof File) {
                            allFiles.push(rootFiles[r].name);
                        }
                    }
                }
                
                // Sort alphabetically like Pandoc
                allFiles.sort();
                
                return allFiles;
                
            } catch (e) {
                return [];
            }
        };
        
        /**
         * Classifies templates by alphabetical order relative to core templates
         * @param {Folder} templateFolder - Template folder
         * @return {Object} Classified templates
         * @private
         */
        this._classifyTemplatesByOrder = function(templateFolder) {
            var templates = {
                before: [],
                frontmatter: null,
                bodymatter: null,
                backmatter: null,
                after: [],
                cover: null,
                specialized: [] // Nouveau: templates spécialisés
            };
            
            if (!templateFolder || !templateFolder.exists) return templates;
            
            var files = templateFolder.getFiles("*.indd");
            var sortedFiles = [];
            
            // Convert to array and sort alphabetically
            for (var i = 0; i < files.length; i++) {
                sortedFiles.push(files[i]);
            }
            sortedFiles.sort(function(a, b) {
                return a.name.toLowerCase() < b.name.toLowerCase() ? -1 : 1;
            });
            
            var frontmatterFound = false;
            var bodymatterFound = false;
            var backmatterFound = false;
            
            for (var j = 0; j < sortedFiles.length; j++) {
                var file = sortedFiles[j];
                var filename = file.name.toLowerCase();
                
                if (filename.indexOf('cover') !== -1) {
                    templates.cover = file;
                } else if (filename.indexOf('frontmatter') !== -1) {
                    templates.frontmatter = file;
                    frontmatterFound = true;
                } else if (filename.indexOf('bodymatter') !== -1) {
                    templates.bodymatter = file;
                    bodymatterFound = true;
                } else if (filename.indexOf('backmatter') !== -1) {
                    templates.backmatter = file;
                    backmatterFound = true;
                } else {
                    // Check if it's a specialized template (contains specific keywords)
                    var isSpecialized = this._isSpecializedTemplate(filename);
                    
                    if (isSpecialized) {
                        templates.specialized.push(file);
                    } else {
                        // Classify as before/after based on alphabetical position
                        if (!frontmatterFound) {
                            templates.before.push(file);
                        } else if (bodymatterFound || backmatterFound) {
                            templates.after.push(file);
                        } else {
                            // Between frontmatter and backmatter, add to before
                            templates.before.push(file);
                        }
                    }
                }
            }
            
            return templates;
        };
        
        /**
         * Checks if a template is specialized (contains specific content keywords)
         * @param {string} filename - Template filename
         * @return {boolean} True if template is specialized
         * @private
         */
        this._isSpecializedTemplate = function(filename) {
            var specializedKeywords = [
                'bibliographie', 'bibliography', 'biblio',
                'index', 'indexes', 'indices',
                'annexe', 'appendix', 'appendices',
                'glossaire', 'glossary',
                'preface', 'préface',
                'conclusion', 'epilogue', 'épilogue',
                'remerciements', 'acknowledgments',
                'dedication', 'dédicace',
                'colophon',
                'notes'
            ];
            
            for (var k = 0; k < specializedKeywords.length; k++) {
                if (filename.indexOf(specializedKeywords[k]) !== -1) {
                    return true;
                }
            }
            
            return false;
        };
        
        /**
         * Determines template type based on Markdown filename content
         * @param {string} mdFileName - Markdown filename
         * @return {string} Template type: 'frontmatter', 'bodymatter', or 'backmatter'
         * @private
         */
        this._determineTemplateType = function(mdFileName) {
            var filename = mdFileName.toLowerCase();
            
            // Frontmatter keywords
            var frontmatterKeywords = [
                'prologue', 'abstract', 'acknowledgments', 'copyright', 'dedication', 
                'credits', 'imprint', 'foreword', 'preface', 'introduction',
                'avant-propos', 'préface', 'dédicace', 'remerciements', 'crédits'
            ];
            
            // Backmatter keywords
            var backmatterKeywords = [
                'appendix', 'colophon', 'bibliography', 'index', 'conclusion',
                'annexe', 'bibliographie', 'glossaire', 'épilogue', 'postface'
            ];
            
            // Bodymatter keywords  
            var bodymatterKeywords = [
                'chapter', 'chapitre', 'part', 'partie'
            ];
            
            // Check each category
            for (var i = 0; i < frontmatterKeywords.length; i++) {
                if (filename.indexOf(frontmatterKeywords[i]) !== -1) {
                    return 'frontmatter';
                }
            }
            
            for (var j = 0; j < backmatterKeywords.length; j++) {
                if (filename.indexOf(backmatterKeywords[j]) !== -1) {
                    return 'backmatter';
                }
            }
            
            for (var k = 0; k < bodymatterKeywords.length; k++) {
                if (filename.indexOf(bodymatterKeywords[k]) !== -1) {
                    return 'bodymatter';
                }
            }
            
            // Default fallback
            return 'bodymatter';
        };
        
        /**
         * Selects the best template for a Markdown file with priority matching
         * @param {string} mdFileName - Markdown filename
         * @param {string} templateType - Detected template type
         * @return {File|null} Best matching template
         * @private
         */
        this._selectBestTemplate = function(mdFileName, templateType) {
            if (!mdFileName) return this.templates.bodymatter;
            
            // Extract keywords from MD filename (remove numbers and extension)
            var mdKeywords = this._extractKeywords(mdFileName);
            
            // Build list of all available templates with their keywords
            var allTemplates = [];
            
            // Add templates from all categories
            if (this.templates.frontmatter) {
                allTemplates.push({
                    file: this.templates.frontmatter,
                    type: 'frontmatter',
                    keywords: this._extractKeywords(this.templates.frontmatter.name)
                });
            }
            
            if (this.templates.bodymatter) {
                allTemplates.push({
                    file: this.templates.bodymatter,
                    type: 'bodymatter', 
                    keywords: this._extractKeywords(this.templates.bodymatter.name)
                });
            }
            
            if (this.templates.backmatter) {
                allTemplates.push({
                    file: this.templates.backmatter,
                    type: 'backmatter',
                    keywords: this._extractKeywords(this.templates.backmatter.name)
                });
            }
            
            // Add all before templates
            for (var b = 0; b < this.templates.before.length; b++) {
                allTemplates.push({
                    file: this.templates.before[b],
                    type: 'before',
                    keywords: this._extractKeywords(this.templates.before[b].name)
                });
            }
            
            // Add all specialized templates
            for (var s = 0; s < this.templates.specialized.length; s++) {
                allTemplates.push({
                    file: this.templates.specialized[s],
                    type: 'specialized',
                    keywords: this._extractKeywords(this.templates.specialized[s].name)
                });
            }
            
            // Find best match based on keyword similarity
            var bestMatch = null;
            var bestScore = 0;
            
            for (var t = 0; t < allTemplates.length; t++) {
                var template = allTemplates[t];
                var score = this._calculateMatchScore(mdKeywords, template.keywords, template.type, templateType);
                
                if (score > bestScore) {
                    bestScore = score;
                    bestMatch = template.file;
                }
            }
            
            // Si on a trouvé un match spécialisé avec un bon score, l'utiliser
            if (bestMatch && bestScore >= 60) {
                return bestMatch;
            }
            
            // Sinon, utiliser la logique de fallback par type
            if (templateType === 'frontmatter' && this.templates.frontmatter) {
                return this.templates.frontmatter;
            } else if (templateType === 'backmatter' && this.templates.backmatter) {
                return this.templates.backmatter;
            } else {
                return this.templates.bodymatter;
            }
        };
        
        /**
         * Extracts meaningful keywords from a filename
         * @param {string} filename - Filename to process
         * @return {Array} Array of keywords
         * @private
         */
        this._extractKeywords = function(filename) {
            if (!filename) return [];
            
            // Remove file extension and common prefixes
            var clean = filename.toLowerCase()
                .replace(/\.indd$/, '')
                .replace(/\.md$/, '')
                .replace(/^[a-z]+-\d+-\d*-?/, '') // Remove patterns like "LIVRE-5-2-"
                .replace(/^\d+-/, ''); // Remove patterns like "13-"
            
            // Split by common separators and filter out short words
            var allWords = clean.split(/[-_\s]+/);
            var words = [];
            for (var w = 0; w < allWords.length; w++) {
                if (allWords[w].length > 2) { // Keep words longer than 2 characters
                    words.push(allWords[w]);
                }
            }
            
            return words;
        };
        
        /**
         * Calculates match score between MD keywords and template keywords
         * @param {Array} mdKeywords - Keywords from Markdown filename
         * @param {Array} templateKeywords - Keywords from template filename  
         * @param {string} templateType - Template type category
         * @param {string} expectedType - Expected template type for MD file
         * @return {number} Match score (0-100)
         * @private
         */
        this._calculateMatchScore = function(mdKeywords, templateKeywords, templateType, expectedType) {
            var score = 0;
            
            // Base score for type matching
            if (templateType === expectedType) {
                score += 30;
            } else if (templateType === 'specialized') {
                // Specialized templates get a bonus for keyword matching
                score += 20;
            }
            
            // Keyword matching score
            var keywordScore = 0;
            var totalPossible = Math.max(mdKeywords.length, templateKeywords.length);
            
            if (totalPossible > 0) {
                for (var i = 0; i < mdKeywords.length; i++) {
                    var mdKeyword = mdKeywords[i];
                    
                    for (var j = 0; j < templateKeywords.length; j++) {
                        var templateKeyword = templateKeywords[j];
                        
                        // Exact match
                        if (mdKeyword === templateKeyword) {
                            keywordScore += 60;
                        }
                        // Partial match (one contains the other)
                        else if (mdKeyword.indexOf(templateKeyword) !== -1 || 
                                 templateKeyword.indexOf(mdKeyword) !== -1) {
                            keywordScore += 30;
                        }
                        // Similarity for common terms
                        else if (this._areSimilarTerms(mdKeyword, templateKeyword)) {
                            keywordScore += 15;
                        }
                    }
                }
                
                // Normalize keyword score
                keywordScore = Math.min(70, keywordScore); // Cap at 70 points
            }
            
            return score + keywordScore;
        };
        
        /**
         * Checks if two terms are semantically similar
         * @param {string} term1 - First term
         * @param {string} term2 - Second term
         * @return {boolean} True if terms are similar
         * @private
         */
        this._areSimilarTerms = function(term1, term2) {
            var synonyms = {
                'bibliographie': ['bibliography', 'biblio', 'references'],
                'bibliography': ['bibliographie', 'biblio', 'references'],
                'index': ['indexes', 'indices'],
                'annexe': ['appendix', 'appendices'],
                'appendix': ['annexe', 'appendices'],
                'glossaire': ['glossary'],
                'glossary': ['glossaire'],
                'preface': ['préface', 'avant-propos'],
                'préface': ['preface', 'avant-propos'],
                'conclusion': ['epilogue', 'épilogue'],
                'epilogue': ['conclusion', 'épilogue'],
                'épilogue': ['conclusion', 'epilogue']
            };
            
            var term1Synonyms = synonyms[term1] || [];
            var term2Synonyms = synonyms[term2] || [];
            
            return arrayContains(term1Synonyms, term2) || arrayContains(term2Synonyms, term1);
        };
    }
    
    /**
     * @namespace LogManager
     * @description Centralized error logging system
     */
    var LogManager = {
        /**
         * Log an error message with optional Error object
         * @param {string} message - Error message
         * @param {Error} [error] - Optional Error object
         */
        logError: function(message, error) {
            if (error) {
                alert(message + ": " + error.message + (error.line ? "\nLine: " + error.line : ""));
            } else {
                alert(message);
            }
            
            // Future: Could log to file or console
        },
        
        /**
         * Log an informational message
         * @param {string} message - Info message
         */
        logInfo: function(message) {
            // For future implementation, currently silent
        }
    };
    
    /**
     * Auto-detects Build/InDesign folder for book generation
     * @param {string} yamlPath - Path to the YAML file
     * @return {Folder|null} Found Build/InDesign folder or null
     */
    function autoDetectBuildFolder(yamlPath) {
        if (!yamlPath) return null;
        
        var yamlFile = File(yamlPath);
        if (!yamlFile.exists) return null;
        
        var searchDir = yamlFile.parent;
        
        // Define possible build folder patterns
        var buildVariants = [
            ["Build", "InDesign"], ["build", "indesign"], 
            ["Build", "indesign"], ["build", "InDesign"],
            ["output", "InDesign"], ["Output", "InDesign"],
            ["export", "InDesign"], ["Export", "InDesign"]
        ];
        
        // Search in current directory and parent
        var searchDirs = [searchDir, searchDir.parent];
        
        for (var d = 0; d < searchDirs.length; d++) {
            var baseDir = searchDirs[d];
            if (!baseDir || !baseDir.exists) continue;
            
            // Try each build folder combination
            for (var v = 0; v < buildVariants.length; v++) {
                var pathParts = buildVariants[v];
                var testPath = baseDir.fsName;
                
                for (var p = 0; p < pathParts.length; p++) {
                    testPath += "/" + pathParts[p];
                }
                
                var testFolder = new Folder(testPath);
                if (testFolder.exists) {
                    return testFolder;
                }
            }
            
            // Also check for simple "Build" folder
            var simpleBuildPath = baseDir.fsName + "/Build";
            var simpleBuildFolder = new Folder(simpleBuildPath);
            if (simpleBuildFolder.exists) {
                return simpleBuildFolder;
            }
        }
        
        return null;
    }
    
    /**
     * Auto-detects template folder based on YAML file location
     * @param {string} yamlPath - Path to the YAML file
     * @return {Folder|null} Found template folder or null
     */
    function autoDetectTemplateFolder(yamlPath) {
        if (!yamlPath) return null;
        
        var yamlFile = File(yamlPath);
        if (!yamlFile.exists) return null;
        
        var searchDir = yamlFile.parent;
        
        // Define possible template folder patterns
        var templateVariants = [
            "template", "Template", "TEMPLATE", "templates", "Templates", "TEMPLATES",
            "indesign", "InDesign", "INDESIGN"
        ];
        
        // Define possible subdirectory combinations
        var subdirCombinations = [
            ["template"], ["Template"], ["templates"], ["Templates"],
            ["template", "indesign"], ["Template", "InDesign"], 
            ["templates", "indesign"], ["Templates", "InDesign"],
            ["indesign"], ["InDesign"]
        ];
        
        // Search in current directory, parent, and grandparent
        var searchDirs = [searchDir, searchDir.parent, searchDir.parent ? searchDir.parent.parent : null];
        
        for (var d = 0; d < searchDirs.length; d++) {
            var baseDir = searchDirs[d];
            if (!baseDir || !baseDir.exists) continue;
            
            // Try each subdirectory combination
            for (var c = 0; c < subdirCombinations.length; c++) {
                var pathParts = subdirCombinations[c];
                var testPath = baseDir.fsName;
                
                for (var p = 0; p < pathParts.length; p++) {
                    testPath += "/" + pathParts[p];
                }
                
                var testFolder = new Folder(testPath);
                if (testFolder.exists) {
                    // Check if folder contains .indd files
                    var inddFiles = testFolder.getFiles("*.indd");
                    if (inddFiles.length > 0) {
                        return testFolder;
                    }
                }
            }
        }
        
        return null;
    }
    
    /**
     * @namespace BookUtils
     * @description Utilities for book creation and manipulation
     */
    var BookUtils = {
        /**
         * @namespace ISBN
         * @description ISBN validation and EAN13 barcode generation
         */
        ISBN: {
            /**
             * Validates an ISBN input
             * @param {string} isbnInput - Raw ISBN string
             * @return {Object} Validation result with status and message
             */
            validate: function(isbnInput) {
                // Check for placeholder pattern (978-2-940426-XX-X or similar)
                var placeholderPattern = /^978-[\d-]*[X-]+[\d-]*$/i;
                if (placeholderPattern.test(isbnInput)) {
                    return { 
                        valid: true, 
                        result: isbnInput,
                        message: "Placeholder ISBN accepted" 
                    };
                }
                
                var isbnDigits = isbnInput.replace(/\D/g, "");
                
                if (isbnDigits.length < 12) {
                    return { valid: false, message: "An ISBN must contain at least 12 digits." };
                }
                
                if (isbnDigits.length === 12) {
                    var checkDigit = this.calculateCheckDigit(isbnDigits);
                    isbnDigits += checkDigit;
                    return { 
                        valid: true, 
                        result: isbnDigits,
                        message: "Check digit added: " + checkDigit
                    };
                }
                
                if (isbnDigits.length === 13) {
                    if (!this.isEAN13Valid(isbnDigits)) {
                        var correctDigit = this.calculateCheckDigit(isbnDigits.substring(0, 12));
                        return { 
                            valid: false, 
                            message: "Invalid check digit. It should be: " + correctDigit
                        };
                    }
                    return { valid: true, result: isbnDigits };
                }
                
                return { valid: false, message: "Invalid ISBN format." };
            },
            
            /**
             * Calculates EAN13 check digit
             * @param {string} code - 12-digit code
             * @return {number} Check digit (0-9)
             */
            calculateCheckDigit: function(code) {
                var sum = 0;
                for (var i = 0; i < 12; i++) {
                    var digit = parseInt(code.charAt(i), 10);
                    sum += (i % 2 === 0) ? digit : digit * 3;
                }
                return (10 - (sum % 10)) % 10;
            },
            
            /**
             * Validates an EAN13 code
             * @param {string} input - 13-digit code
             * @return {boolean} Is valid EAN13
             */
            isEAN13Valid: function(input) {
                return input.length === 13 && 
                    parseInt(input.charAt(12), 10) === this.calculateCheckDigit(input.slice(0, 12));
            },
            
            /**
             * Encodes EAN13 into binary pattern
             * @param {string} code - 13-digit EAN13 code
             * @return {string} Binary pattern for barcode
             */
            encodeEAN13: function(code) {
                var LEFT_ODD = {0:"0001101",1:"0011001",2:"0010011",3:"0111101",4:"0100011",
                               5:"0110001",6:"0101111",7:"0111011",8:"0110111",9:"0001011"};
                var LEFT_EVEN = {0:"0100111",1:"0110011",2:"0011011",3:"0100001",4:"0011101",
                                5:"0111001",6:"0000101",7:"0010001",8:"0001001",9:"0010111"};
                var RIGHT = {0:"1110010",1:"1100110",2:"1101100",3:"1000010",4:"1011100",
                            5:"1001110",6:"1010000",7:"1000100",8:"1001000",9:"1110100"};
                var PARITY = {0:"OOOOOO",1:"OOEOEE",2:"OOEEOE",3:"OOEEEO",4:"OEOOEE",
                             5:"OEEOOE",6:"OEEEOO",7:"OEOEOE",8:"OEOEEO",9:"OEEOEO"};
            
                var pattern = "101";
                var parity = PARITY[parseInt(code.charAt(0))];
                for (var i = 1; i <= 6; i++) {
                    var digit = parseInt(code.charAt(i));
                    pattern += (parity.charAt(i - 1) === 'O' ? LEFT_ODD[digit] : LEFT_EVEN[digit]);
                }
                pattern += "01010";
                for (var j = 7; j <= 12; j++) pattern += RIGHT[parseInt(code.charAt(j))];
                pattern += "101";
                return pattern;
            },
            
            /**
             * Draws EAN13 barcode in container
             * @param {Rectangle} container - Target container
             * @param {string} code - EAN13 code
             * @param {Document} doc - InDesign document
             */
            drawBarcode: function(container, code, doc) {
                var bounds = container.geometricBounds;
                var width = bounds[3] - bounds[1];
                var height = bounds[2] - bounds[0];
                var moduleWidth = width / 95;
                var black = doc.swatches.itemByName("Black");
                
                var binary = this.encodeEAN13(code);
            
                for (var i = 0; i < binary.length; i++) {
                    if (binary.charAt(i) === '1') {
                        var x1 = bounds[1] + i * moduleWidth;
                        var rect = container.parentPage.rectangles.add();
                        rect.geometricBounds = [bounds[0], x1, bounds[2], x1 + moduleWidth];
                        rect.fillColor = black;
                        rect.strokeColor = doc.swatches.itemByName("None");
                    }
                }
            }
        },
        
        /**
         * @namespace Document
         * @description InDesign document operations
         */
        Document: {
            /**
             * Adds a custom text variable to a document
             * @param {Document} doc - InDesign document
             * @param {string} varName - Variable name
             * @param {string} varContent - Variable content
             */
            addCustomVariable: function(doc, varName, varContent) {
                // Only create InDesign text variables for title and author
                if (varName !== "title" && varName !== "author") {
                    return; // Skip other variables
                }
                
                // Convert variable names to match conventions
                var nameMap = {
                    "author": "Book Author",
                    "title": "Book Title"
                };
                
                var displayName = nameMap[varName] || varName;
                
                // Clean content by removing <br> tags
                var cleanContent = varContent ? varContent.replace(/<br\s*\/?>/gi, "") : "";
                
                try {
                    doc.textVariables.item(displayName).variableOptions.contents = cleanContent;
                } catch(e) {
                    try {
                        doc.textVariables.add({
                            name: displayName, 
                            variableType: VariableTypes.CUSTOM_TEXT_TYPE
                        }).variableOptions.contents = cleanContent;
                    } catch(e) {
                        // Continue silently if variable can't be added
                    }
                }
            },
            
            /**
             * Replaces text placeholders with formatted content
             * @param {Document} doc - InDesign document
             * @param {Object} bookInfo - Book metadata
             * @param {Object} displayOptions - Display settings
             * @return {boolean} Success status
             */
            replaceTextPlaceholders: function(doc, bookInfo, displayOptions) {
                // Prepare placeholder values
                var values = {
                    "<<Book_Author>>": bookInfo.author || "",
                    "<<Book_Title>>": bookInfo.title || "",
                    "<<Subtitle>>": bookInfo.subtitle || "",
                    "<<ISBN_Print>>": bookInfo.isbnPrint || "",
                    "<<ISBN_Ebook>>": bookInfo.isbnEbook || "",
                    "<<Translation>>": bookInfo.translation || "",
                    "<<Critical_Apparatus>>": bookInfo.critical || "",
                    "<<Print_Date>>": bookInfo.printDate || "",
                    "<<Editions>>": bookInfo.editions || "",
                    "<<Funding>>": bookInfo.funding || "",
                    "<<Rights>>": bookInfo.rights || "",
                    "<<Price>>": bookInfo.price || ""
                };
                
                // Add labels for fields with special options
                if (displayOptions && displayOptions.showOriginalTitleLabel && bookInfo.originalTitle) {
                    values["<<Original_Title>>"] = displayOptions.originalTitleLabelText + bookInfo.originalTitle;
                } else {
                    values["<<Original_Title>>"] = bookInfo.originalTitle || "";
                }
                
                if (displayOptions && displayOptions.showCoverCreditLabel && bookInfo.coverCredit) {
                    values["<<Cover_Credit>>"] = displayOptions.coverCreditLabelText + bookInfo.coverCredit;
                } else {
                    values["<<Cover_Credit>>"] = bookInfo.coverCredit || "";
                }
                
                // Placeholders that should be completely removed if empty
                var emptyFields = [
                    "<<Original_Title>>",
                    "<<Translator_Display>>",
                    "<<Critical_Display>>",
                    "<<Cover_Credit>>", 
                    "<<Editions>>", 
                    "<<Funding>>"
                ];
                                  
                // Process each text frame in document
                for (var i = 0; i < doc.textFrames.length; i++) {
                    var tf = doc.textFrames[i];
                    if (!tf.contents) continue;
                    
                    var originalContent = tf.contents;
                    var newContent = originalContent;
                    var hasPlaceholder = false;
                    
                    // For each placeholder
                    for (var placeholder in values) {
                        if (newContent.indexOf(placeholder) !== -1) {
                            // Found a placeholder
                            hasPlaceholder = true;
                            var value = values[placeholder];
                            
                            if (!value && arrayContains(emptyFields, placeholder)) {
                                // Case 1: Empty field that should be removed with its line
                                // Temporarily mark for removal
                                newContent = newContent.replace(placeholder, "###EMPTY_PLACEHOLDER###");
                            } else {
                                // Case 2: Replace with value (may be empty)
                                newContent = newContent.replace(placeholder, value);
                            }
                        }
                    }
                    
                    // Remove lines containing empty placeholders
                    if (hasPlaceholder && newContent.indexOf("###EMPTY_PLACEHOLDER###") !== -1) {
                        var lines = newContent.split(/\r|\n/);
                        var filteredLines = [];
                        
                        for (var j = 0; j < lines.length; j++) {
                            if (lines[j].indexOf("###EMPTY_PLACEHOLDER###") === -1) {
                                filteredLines.push(lines[j]);
                            }
                        }
                        
                        newContent = filteredLines.join("\r");
                    }
                    
                    // If content was modified, apply basic formatting
                    if (hasPlaceholder && originalContent !== newContent) {
                        TextUtils.applyFormattedText(tf, newContent, doc);
                    }
                }
                
                return true;
            },
            
            /**
             * Replaces EAN13 placeholders with barcodes
             * @param {Document} doc - InDesign document
             * @param {string} isbnPrint - Print ISBN
             * @param {string} isbnEbook - Ebook ISBN
             * @return {boolean} Success status
             */
            replaceEAN13Placeholders: function(doc, isbnPrint, isbnEbook) {
                var placeholders = {
                    "<<EAN13_Print>>": isbnPrint,
                    "<<EAN13_Ebook>>": isbnEbook
                };
                
                for (var placeholder in placeholders) {
                    var isbnValue = placeholders[placeholder];
                    if (!isbnValue) continue;
                    
                    // Clean ISBN to get only digits
                    var isbnDigits = isbnValue.replace(/\D/g, "");
                    
                    // Ensure it's a valid EAN13
                    if (isbnDigits.length === 12) {
                        isbnDigits += BookUtils.ISBN.calculateCheckDigit(isbnDigits);
                    }
                    
                    if (!BookUtils.ISBN.isEAN13Valid(isbnDigits)) {
                        continue; // Skip invalid ISBNs
                    }
                    
                    // Find placeholders 
                    for (var i = 0; i < doc.textFrames.length; i++) {
                        var tf = doc.textFrames[i];
                        if (tf.contents && tf.contents.replace(/\s+/g, '') === placeholder) {
                            try {
                                var bounds = tf.geometricBounds;
                                var page = tf.parentPage;
                                tf.remove();
                                
                                var container = page.rectangles.add({geometricBounds: bounds});
                                container.fillColor = doc.swatches.itemByName("None");
                                container.strokeWeight = 0;
                                container.strokeColor = doc.swatches.itemByName("None");
                                
                                BookUtils.ISBN.drawBarcode(container, isbnDigits, doc);
                            } catch (e) {
                                LogManager.logError("Error replacing " + placeholder, e);
                            }
                        }
                    }
                }
                return true;
            }
        },
        
        /**
         * @namespace File
         * @description File operations for YAML and Markdown
         */
        File: {
            /**
             * Imports YAML metadata
             * @param {File} yamlFile - YAML file
             * @return {Object} Import result with metadata
             */
            importYAML: function(yamlFile) {
                if (!yamlFile || !yamlFile.exists) {
                    throw new Error("YAML file does not exist.");
                }
                
                try {
                    yamlFile.open("r");
                    var content = yamlFile.read();
                    yamlFile.close();
                    
                    var yamlData = YAMLParser.parse(content);
                    
                    // Use PandocMapper to handle complex structures
                    var result = PandocMapper.mapToBookCreator(yamlData);
                    
                    return {
                        result: result,
                        yamlMeta: yamlData,
                        yamlPath: yamlFile.fsName
                    };
                } catch (e) {
                    if (yamlFile.open) yamlFile.close();
                    throw new Error("Error importing YAML: " + e.message);
                }
            },
            
            /**
             * Exports book metadata to YAML
             * @param {Object} bookInfo - Book metadata
             * @param {File} yamlFile - Target YAML file
             * @return {boolean} Success status
             */
            exportYAML: function(bookInfo, yamlFile) {
                if (!yamlFile) {
                    throw new Error("No file specified for YAML export.");
                }
                
                try {
                    // Mapping between BookCreator fields and Pandoc fields
                    var fieldMappings = {
                        "title": "title",
                        "subtitle": "subtitle",
                        "author": "author", 
                        "printDate": "published-print",
                        "language": "lang",
                        "rights": "rights",
                        
                        // ISBN specific fields
                        "isbnPrint": "isbn-print",
                        "isbnEbook": "isbn-ebook",
                        
                        // BookCreator specific fields preserved as-is
                        "translator-display": "translation",
                        "critical-display": "critical",
                        "originalTitle": "originalTitle", 
                        "cover-note": "coverCredit",
                        "editions": "editions",
                        "funding": "funding",
                        "price": "price"
                    };
                    
                    var yamlData = {};
                    
                    // Create YAML object with correct fields
                    for (var srcField in fieldMappings) {
                        if (bookInfo.hasOwnProperty(srcField) && bookInfo[srcField]) {
                            var destField = fieldMappings[srcField];
                            yamlData[destField] = bookInfo[srcField];
                        }
                    }
                    
                    // Export title as complex structure if both title and subtitle exist
                    if (bookInfo.title || bookInfo.subtitle) {
                        yamlData.title = [];
                        if (bookInfo.title) {
                            yamlData.title.push({
                                type: "main",
                                text: bookInfo.title
                            });
                        }
                        if (bookInfo.subtitle) {
                            yamlData.title.push({
                                type: "subtitle", 
                                text: bookInfo.subtitle
                            });
                        }
                        // Remove simple fields since we're using complex structure
                        delete yamlData.title; // Remove if exists from simple mapping
                        delete yamlData.subtitle; // Remove if exists from simple mapping
                    }
                    
                    // Export creator as complex structure
                    if (bookInfo.author) {
                        yamlData.creator = [{
                            role: "author",
                            text: bookInfo.author
                        }];
                        // Remove simple field
                        delete yamlData.author; // Remove if exists from simple mapping
                    }
                    
                    // Export ISBN with different formats: complex for ebook, simple for print
                    if (bookInfo.isbnEbook) {
                        // Ebook uses complex identifier structure
                        yamlData.identifier = [{
                            scheme: "ISBN",
                            text: bookInfo.isbnEbook
                        }];
                    }
                    
                    if (bookInfo.isbnPrint) {
                        // Print uses simple isbn-print field (handled by normal mappings)
                        // Nothing special needed here, normal mapping will handle it
                    }
                    
                    // Don't delete isbn-print since we want to keep it simple
                    delete yamlData["isbn-ebook"];
                    
                    // Include other unmapped fields
                    for (var field in bookInfo) {
                        if (!fieldMappings.hasOwnProperty(field) && bookInfo[field]) {
                            yamlData[field] = bookInfo[field];
                        }
                    }
                    
                    // Generate YAML string
                    var yamlContent = "---\n";  // YAML front matter starts with ---
                    yamlContent += YAMLParser.stringify(yamlData);
                    yamlContent += "---\n";  // YAML front matter ends with ---
                    
                    yamlFile.encoding = "UTF-8";
                    yamlFile.open("w");
                    yamlFile.write(yamlContent);
                    yamlFile.close();
                    
                    return true;
                } catch (e) {
                    if (yamlFile.open) yamlFile.close();
                    throw new Error("Error exporting YAML: " + e.message);
                }
            },
            
            /**
             * Detects Markdown elements in YAML data
             * @param {Object} yamlData - Parsed YAML data
             * @return {Object} Detection result
             */
            detectMarkdownElements: function(yamlData) {
                var result = {
                    hasInputFiles: false
                };
                
                // Plus besoin de input-files, on considère qu'il y a des fichiers si le YAML existe
                result.hasInputFiles = true;
                
                return result;
            }
        }
    };
    
    /**
     * @namespace UI
     * @description User interface components
     */
    var UI = {
        /**
         * Creates the main application window
         * @return {Window} InDesign dialog window
         */
        createMainWindow: function() {
            var book = new Book("", {});
            
            var win = new Window('dialog', I18n.__('title', VERSION));
            win.alignChildren = 'fill';
            
            // Language selector
            var langGroup = win.add('group');
            langGroup.orientation = "row";
            langGroup.alignment = "right";
            
            // Ajouter le texte d'attribution
            langGroup.add("statictext", undefined, "entremonde / Spectral lab");
            // Ajouter un petit espace
            langGroup.add("statictext", undefined, "  ");
            // Ajouter le dropdown comme avant
            var langDropdown = langGroup.add('dropdownlist', undefined, ['En', 'Fr']);
            
            // Select current language
            langDropdown.selection = I18n.getLanguage() === 'fr' ? 1 : 0;
            
            langDropdown.onChange = function() {
                // Change language
                I18n.setLanguage(langDropdown.selection.index === 1 ? 'fr' : 'en');
                
                // Close and reopen the window to apply changes immediately
                var currentLanguage = I18n.getLanguage();
                win.close();
                var newWindow = UI.createMainWindow();
                newWindow.show();
            };
            
            // Book Information button
            var infoBtn = win.add('button', undefined, I18n.__('bookInformation'));
            infoBtn.onClick = function() {
                UI.createInfoWindow(book.info, book.displayOptions, book.markdownOptions, updateMdOptions, bookNameInput, book, templateFolderText, templateDetailsText);
            };
            
            // Main fields
            win.add('statictext', undefined, I18n.__('fileNamePrefix'));
            var bookNameInput = win.add('edittext', undefined, '');
            
            // Templates panel
            var templatesPanel = win.add('panel', undefined, 'Templates Folder');
            templatesPanel.alignChildren = 'fill';
            templatesPanel.preferredSize.height = 120; // Taille fixe
            
            // Template folder selection
            var templateFolderBtn = templatesPanel.add('button', undefined, 'Select Templates Folder');
            var templateFolderText = templatesPanel.add('statictext', undefined, 'No folder selected.');
            var templateDetailsText = templatesPanel.add('statictext', undefined, '');
            templateDetailsText.preferredSize.height = 60;
            
            templateFolderBtn.onClick = function() {
                var folder = Folder.selectDialog('Choose Templates Folder');
                if (folder) {
                    book.templateFolder = folder;
                    book.templates = book._classifyTemplatesByOrder(folder);
                    
                    templateFolderText.text = folder.name;
                    
                    // Show discovered templates
                    var details = "Discovered:\n";
                    details += "Before: " + book.templates.before.length + " files\n";
                    details += "Frontmatter: " + (book.templates.frontmatter ? "✓" : "✗") + "\n";
                    details += "Bodymatter: " + (book.templates.bodymatter ? "✓" : "✗") + "\n";
                    details += "Backmatter: " + (book.templates.backmatter ? "✓" : "✗") + "\n";
                    details += "After: " + book.templates.after.length + " files\n";
                    details += "Cover: " + (book.templates.cover ? "✓" : "✗");
                    
                    templateDetailsText.text = details;
                }
            };
            
            // Text folder panel (optional manual selection)
            var textPanel = win.add('panel', undefined, 'Text Folder');
            textPanel.alignChildren = 'fill';
            textPanel.preferredSize.height = 120; // Même taille que Templates
            
            var textFolderBtn = textPanel.add('button', undefined, 'Select Text Folder (optional)');
            
            // Markdown detection status
            var mdDiagnostic = textPanel.add('statictext', undefined, I18n.__('noYAMLLoaded'));
            mdDiagnostic.preferredSize.height = 60;
            
            textFolderBtn.onClick = function() {
                var folder = Folder.selectDialog('Choose Text Folder');
                if (folder) {
                    book.markdownOptions.textFolderPath = folder.fsName;
                }
            };
            
            function updateMdOptions() {
                // Auto-enable markdown injection if files are detected
                book.markdownOptions.injectMarkdown = book.markdownOptions.hasInputFiles;
                
                // Update diagnostic text in main window
                if (typeof mdDiagnostic !== 'undefined') {
                    if (book.markdownOptions.yamlPath) {
                        if (book.markdownOptions.hasInputFiles) {
                            var allMdFiles = book._getAllMarkdownFilesFromFolder();
                            var fileCount = allMdFiles.length;
                            mdDiagnostic.text = I18n.__('inputFilesDetected') + " (" + fileCount + " files)";
                        } else {
                            mdDiagnostic.text = I18n.__('inputFilesNotDetected');
                        }
                    } else {
                        mdDiagnostic.text = I18n.__('noYAMLLoaded');
                    }
                }
            }
            
            // Action buttons
            var actionBtns = win.add('group');
            actionBtns.alignment = 'right';
            actionBtns.add('button', undefined, I18n.__('cancel'), {name:'cancel'});
            var createBtn = actionBtns.add('button', undefined, I18n.__('createBook'), {name:'ok'});
            
            createBtn.onClick = function() {
                // Update book info
                book.name = bookNameInput.text;
                
                // Validate data
                var validation = book.validate();
                if (!validation.valid) {
                    alert(validation.message);
                    return;
                }
                
                // Generate book - with auto-detected default folder
                var defaultFolder = null;
                if (book.markdownOptions && book.markdownOptions.yamlPath) {
                    defaultFolder = autoDetectBuildFolder(book.markdownOptions.yamlPath);
                }
                
                var folder = Folder.selectDialog(I18n.__('chooseDestionationFolder'), defaultFolder);
                if (folder) {
                    try {
                        // Disable InDesign native error messages during execution
                        var originalUserInteractionLevel = app.scriptPreferences.userInteractionLevel;
                        app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;
                        
                        var result = book.generate(folder, win);
                        
                        // Restore user interaction level BEFORE opening the book
                        app.scriptPreferences.userInteractionLevel = originalUserInteractionLevel;
                        
                        if (result) {
                            alert(I18n.__('bookGenerated'));
                            // La fenêtre est déjà fermée dans generate()
                        }
                    } catch (e) {
                        // Restore user interaction level in case of error
                        app.scriptPreferences.userInteractionLevel = originalUserInteractionLevel;
                        LogManager.logError(I18n.__('error'), e);
                    }
                }
            };
            
            return win;
        },
        
        /**
         * Creates the book information window
         * @param {Object} bookInfo - Book metadata
         * @param {Object} displayOptions - Display settings
         * @param {Object} markdownOptions - Markdown settings
         * @param {Function} updateMdDiagnosticCallback - Callback for updating Markdown options
         */
        createInfoWindow: function(bookInfo, displayOptions, markdownOptions, updateMdDiagnosticCallback, bookNameInput, book, templateFolderText, templateDetailsText) {
            var infoWin = new Window("dialog", I18n.__('bookInformation'));
            infoWin.orientation = "column";
            infoWin.alignChildren = "fill";
            infoWin.spacing = 10;
            
            var inputGroup = infoWin.add("group");
            inputGroup.orientation = "row";
            inputGroup.alignChildren = "top";
            inputGroup.spacing = 20;
            
            // Column 1
            var col1 = inputGroup.add("group");
            col1.orientation = "column";
            col1.alignChildren = "left";
            col1.spacing = 5;
            
            col1.add("statictext", undefined, I18n.__('bookAuthor'));
            var authorInput = col1.add("edittext", undefined, bookInfo.author || "");
            authorInput.characters = 25;
            
            col1.add("statictext", undefined, I18n.__('bookTitle'));
            var titleInput = col1.add("edittext", undefined, bookInfo.title || "");
            titleInput.characters = 25;
            
            col1.add("statictext", undefined, I18n.__('subtitle'));
            var subtitleInput = col1.add("edittext", undefined, bookInfo.subtitle || "");
            subtitleInput.characters = 25;
            
            col1.add("statictext", undefined, I18n.__('isbnPrint'));
            var isbnPrintInput = col1.add("edittext", undefined, bookInfo.isbnPrint || "");
            isbnPrintInput.characters = 25;
            
            col1.add("statictext", undefined, I18n.__('isbnEbook'));
            var isbnEbookInput = col1.add("edittext", undefined, bookInfo.isbnEbook || "");
            isbnEbookInput.characters = 25;
            
            col1.add("statictext", undefined, I18n.__('price'));
            var priceInput = col1.add("edittext", undefined, bookInfo.price || "");
            priceInput.characters = 25;
            
            // Original title with checkbox
            var originalTitleGroup = col1.add("group");
            originalTitleGroup.orientation = "column";
            originalTitleGroup.alignChildren = "left";
            
            originalTitleGroup.add("statictext", undefined, I18n.__('originalTitle'));
            var originalTitleInput = originalTitleGroup.add("edittext", undefined, bookInfo.originalTitle || "");
            originalTitleInput.characters = 25;
            
            // Add prefix
            var originalTitleLabelGroup = originalTitleGroup.add("group");
            originalTitleLabelGroup.orientation = "row";
            originalTitleLabelGroup.alignChildren = "center";
            
            var showOriginalTitleLabel = originalTitleLabelGroup.add("checkbox", undefined, I18n.__('prefix'));
            showOriginalTitleLabel.value = displayOptions ? displayOptions.showOriginalTitleLabel : true;
            
            var originalTitleLabelText = originalTitleLabelGroup.add("edittext", undefined, 
                                        displayOptions.originalTitleLabelText || I18n.__('originalTitlePrefix'));
            originalTitleLabelText.characters = 15;
            
            // Move Print date after Original title
            col1.add("statictext", undefined, I18n.__('printDate'));
            var printDateInput = col1.add("edittext", undefined, bookInfo.printDate || "");
            printDateInput.characters = 25;
            
            // Column 2
            var col2 = inputGroup.add("group");
            col2.orientation = "column";
            col2.alignChildren = "left";
            col2.spacing = 5;
            
            col2.add("statictext", undefined, I18n.__('critical'));
            var criticalInput = col2.add("edittext", undefined, bookInfo.critical || "", {multiline: true});
            criticalInput.preferredSize = [250, 60];
            
            col2.add("statictext", undefined, I18n.__('translation'));
            var translationInput = col2.add("edittext", undefined, bookInfo.translation || "", {multiline: true});
            translationInput.preferredSize = [250, 60];
            
            // Cover credit with checkbox
            var coverCreditGroup = col2.add("group");
            coverCreditGroup.orientation = "column";
            coverCreditGroup.alignChildren = "left";
            
            coverCreditGroup.add("statictext", undefined, I18n.__('coverCredit'));
            var coverCreditInput = coverCreditGroup.add("edittext", undefined, bookInfo.coverCredit || "", {multiline: true});
            coverCreditInput.preferredSize = [250, 40];
            
            // Add prefix
            var coverCreditLabelGroup = coverCreditGroup.add("group");
            coverCreditLabelGroup.orientation = "row";
            coverCreditLabelGroup.alignChildren = "center";
            
            var showCoverCreditLabel = coverCreditLabelGroup.add("checkbox", undefined, I18n.__('prefix'));
            showCoverCreditLabel.value = displayOptions ? displayOptions.showCoverCreditLabel : true;
            
            var coverCreditLabelText = coverCreditLabelGroup.add("edittext", undefined, 
                                      displayOptions.coverCreditLabelText || I18n.__('coverCreditPrefix'));
            coverCreditLabelText.characters = 15;
            
            col2.add("statictext", undefined, I18n.__('editions'));
            var editionsInput = col2.add("edittext", undefined, bookInfo.editions || "", {multiline: true});
            editionsInput.preferredSize = [250, 40];
            
            col2.add("statictext", undefined, I18n.__('funding'));
            var fundingInput = col2.add("edittext", undefined, bookInfo.funding || "", {multiline: true});
            fundingInput.preferredSize = [250, 40];
            
            // Rights field
            col2.add("statictext", undefined, I18n.__('rights'));
            var rightsInput = col2.add("edittext", undefined, bookInfo.rights || "", {multiline: true});
            rightsInput.preferredSize = [250, 40];
            
            // Button group
            var btnGroup = infoWin.add("group");
            btnGroup.orientation = "row";
            btnGroup.alignment = "center";
            btnGroup.spacing = 10;
            
            // YAML import button
            var importBtn = btnGroup.add("button", undefined, I18n.__('importYAML'));
            importBtn.onClick = function() {
                var yamlFile = File.openDialog(I18n.__('importYAML'), "*.yaml;*.yml");
                if (yamlFile) {
                    try {
                        var importResult = BookUtils.File.importYAML(yamlFile);
                        var yamlData = importResult.result;
                        
                        authorInput.text = yamlData.author || "";
                        titleInput.text = yamlData.title || "";
                        subtitleInput.text = yamlData.subtitle || "";
                        isbnPrintInput.text = yamlData.isbnPrint || "";
                        isbnEbookInput.text = yamlData.isbnEbook || "";
                        printDateInput.text = yamlData.printDate || "";
                        originalTitleInput.text = yamlData.originalTitle || "";
                        criticalInput.text = yamlData.critical || "";
                        translationInput.text = yamlData.translation || "";
                        coverCreditInput.text = yamlData.coverCredit || "";
                        editionsInput.text = yamlData.editions || "";
                        fundingInput.text = yamlData.funding || "";
                        rightsInput.text = yamlData.rights || "";
                        priceInput.text = yamlData.price || "";
                        
                        // Auto-remplir le préfixe si vide
                        if (bookNameInput && bookNameInput.text === "" && yamlData.title) {
                            var autoPrefix = normalizeBookTitle(yamlData.title);
                            bookNameInput.text = autoPrefix;
                        }
                        
                        // Store YAML data and check Markdown elements
                        markdownOptions.yamlPath = importResult.yamlPath;
                        markdownOptions.yamlMeta = importResult.yamlMeta;
                        
                        var mdElements = BookUtils.File.detectMarkdownElements(importResult.yamlMeta);
                        markdownOptions.hasInputFiles = mdElements.hasInputFiles;
                        
                        // Auto-detect template folder
                        var detectedTemplateFolder = autoDetectTemplateFolder(importResult.yamlPath);
                        if (detectedTemplateFolder) {
                            // Update the book's template folder and classification
                            book.templateFolder = detectedTemplateFolder;
                            book.templates = book._classifyTemplatesByOrder(detectedTemplateFolder);
                            
                            // Update the main window's template display
                            if (typeof templateFolderText !== 'undefined') {
                                templateFolderText.text = detectedTemplateFolder.name;
                                
                                // Update template details
                                var details = "Discovered:\n";
                                details += "Before: " + book.templates.before.length + " files\n";
                                details += "Frontmatter: " + (book.templates.frontmatter ? "✓" : "✗") + "\n";
                                details += "Bodymatter: " + (book.templates.bodymatter ? "✓" : "✗") + "\n";
                                details += "Backmatter: " + (book.templates.backmatter ? "✓" : "✗") + "\n";
                                details += "After: " + book.templates.after.length + " files\n";
                                details += "Cover: " + (book.templates.cover ? "✓" : "✗");
                                
                                if (typeof templateDetailsText !== 'undefined') {
                                    templateDetailsText.text = details;
                                }
                            }
                        }
                        
                        // Update Markdown options in main window
                        if (updateMdDiagnosticCallback) {
                            updateMdDiagnosticCallback();
                        }
                        
                        // Créer un message complet d'import
                        var importMessage = I18n.__('yamlImportCompleted');
                        
                        // Ajouter l'info sur le dossier template détecté
                        if (detectedTemplateFolder) {
                            var templateInfo = I18n.getLanguage() === 'fr' ? 
                                "\nDossier de templates d\u00E9tect\u00E9: " + detectedTemplateFolder.name :
                                "\nTemplate folder detected: " + detectedTemplateFolder.name;
                            importMessage += templateInfo;
                        } else {
                            var noTemplateInfo = I18n.getLanguage() === 'fr' ? 
                                "\nAucun dossier de templates d\u00E9tect\u00E9 automatiquement." :
                                "\nNo template folder auto-detected.";
                            importMessage += noTemplateInfo;
                        }
                        
                        // Ajouter l'info sur les fichiers Markdown
                        if (markdownOptions.hasInputFiles) {
                            var allMdFiles = book._getAllMarkdownFilesFromFolder();
                            var fileCount = allMdFiles.length;
                            var markdownInfo = I18n.getLanguage() === 'fr' ? 
                                "\nFichiers Markdown d\u00E9tect\u00E9s: " + fileCount + " fichiers" :
                                "\nMarkdown files detected: " + fileCount + " files";
                            importMessage += markdownInfo;
                        } else {
                            var noMarkdownInfo = I18n.getLanguage() === 'fr' ? 
                                "\nAucun fichier Markdown d\u00E9tect\u00E9." :
                                "\nNo Markdown files detected.";
                            importMessage += noMarkdownInfo;
                        }
                        
                        alert(importMessage);
                    } catch (e) {
                        LogManager.logError(I18n.__('error'), e);
                    }
                }
            };
            
            // YAML export button
            var exportBtn = btnGroup.add("button", undefined, I18n.__('exportYAML'));
            exportBtn.onClick = function() {
                var yamlFile = File.saveDialog(I18n.__('exportYAML'), "*.yaml");
                if (yamlFile) {
                    try {
                        var data = {
                            author: authorInput.text,
                            title: titleInput.text,
                            subtitle: subtitleInput.text,
                            isbnPrint: isbnPrintInput.text,
                            isbnEbook: isbnEbookInput.text,
                            critical: criticalInput.text,
                            translation: translationInput.text,
                            printDate: printDateInput.text,
                            originalTitle: originalTitleInput.text,
                            coverCredit: coverCreditInput.text,
                            editions: editionsInput.text,
                            funding: fundingInput.text,
                            rights: rightsInput.text,
                            price: priceInput.text
                        };
                        
                        BookUtils.File.exportYAML(data, yamlFile);
                        alert(I18n.__('yamlExportCompleted'));
                    } catch (e) {
                        LogManager.logError(I18n.__('error'), e);
                    }
                }
            };
            
            // Save button
            var saveInfoBtn = btnGroup.add("button", undefined, I18n.__('save'));
            saveInfoBtn.onClick = function() {
                bookInfo.author = authorInput.text;
                bookInfo.title = titleInput.text;
                bookInfo.subtitle = subtitleInput.text;
                bookInfo.isbnPrint = isbnPrintInput.text;
                bookInfo.isbnEbook = isbnEbookInput.text;
                bookInfo.critical = criticalInput.text;
                bookInfo.translation = translationInput.text;
                bookInfo.printDate = printDateInput.text;
                bookInfo.originalTitle = originalTitleInput.text;
                bookInfo.coverCredit = coverCreditInput.text;
                bookInfo.editions = editionsInput.text;
                bookInfo.funding = fundingInput.text;
                bookInfo.rights = rightsInput.text;
                bookInfo.price = priceInput.text;
                
                // Save display options with new label text fields
                displayOptions.showOriginalTitleLabel = showOriginalTitleLabel.value;
                displayOptions.originalTitleLabelText = originalTitleLabelText.text;
                displayOptions.showCoverCreditLabel = showCoverCreditLabel.value;
                displayOptions.coverCreditLabelText = coverCreditLabelText.text;
                
                infoWin.close();
            };
            
            infoWin.show();
        }
    };
    
    /**
     * Public API for BookCreator
     * @type {Object}
     * @property {Function} init - Initializes and shows the main window
     * @property {string} version - Current version number
     */
    return {
        /**
         * Initializes the BookCreator application
         */
        init: function() {
            var mainWindow = UI.createMainWindow();
            mainWindow.show();
        },
        
        /**
         * Changes the interface language
         * @param {string} lang - Language code ('en' or 'fr')
         */
        setLanguage: function(lang) {
            I18n.setLanguage(lang);
        },
        
        /**
         * Gets the current interface language
         * @return {string} Language code
         */
        getLanguage: function() {
            return I18n.getLanguage();
        },
        
        version: VERSION
    };
})();

// Start the application
BookCreator.init();