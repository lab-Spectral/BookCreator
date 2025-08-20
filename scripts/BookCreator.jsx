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
            this.inList = false;
            this.currentList = [];
            this.inObject = false;
            this.currentObject = {};
            this.lineIndex = 0;
            
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
                        // Add item to current list
                        ctx.currentList.push(ctx.extractListItem(line));
                        line = ctx.getNextLine();
                        continue;
                    } else if (lineIndent < ctx.currentIndent || lineIndent === 0) {
                        // End of list
                        ctx.result[ctx.currentKey] = ctx.currentList;
                        ctx.inList = false;
                        ctx.currentList = [];
                        ctx.currentKey = null;
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
                        ctx.multilineValue += " " + utils.trim(line.replace(/^\s{2}/, ""));
                        line = ctx.getNextLine();
                        continue;
                    } else {
                        // End of multiline value
                        ctx.result[ctx.currentKey] = ctx.multilineValue;
                        ctx.inMultiline = false;
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
                    if (value === "") {
                        // Multiline value starts on next line
                        ctx.inMultiline = true;
                        ctx.multilineValue = "";
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
                'criticalApparatus': 'Critical Apparatus:',
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
                'bookGenerated': 'Book successfully generated!\n\nYou can now open it in InDesign.',
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
                'criticalApparatus': 'Appareil critique :',
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
                'bookGenerated': 'Livre g\u00E9n\u00E9r\u00E9 avec succ\u00E8s !\n\nVous pouvez maintenant l\'ouvrir dans InDesign.',
                'yamlImportCompleted': 'Import YAML termin\u00E9.',
                'yamlExportCompleted': 'Export YAML termin\u00E9.',
                'fileExists': 'Le fichier %s existe d\u00E9j\u00E0. Voulez-vous le remplacer ?',
                'destinationFolderNotExist': 'Le dossier de destination n\'existe pas.',
                'markdownFileNotFound': 'Fichier Markdown non trouv\u00E9 : %s',
                'noTextFrameFound': 'Aucun cadre de texte trouv\u00E9 dans le document.',
                'error': 'Erreur',
                
                // Default prefixes
                'originalTitlePrefix': 'Titre original : ',
                'coverCreditPrefix': 'Couverture : ',
                
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
     * @description Text formatting utilities (simplified without Markdown formatting)
     */
    var TextUtils = {
        /**
         * Applies plain text to a text frame without Markdown formatting
         * @param {TextFrame} textFrame - InDesign text frame to apply text to
         * @param {string} text - Raw text content
         * @param {Document} doc - Parent InDesign document
         */
        applyFormattedText: function(textFrame, text, doc) {
            if (!text) {
                textFrame.contents = "";
                return;
            }
            
            // Process only <br> tags and trailing spaces
            var processedText = text.replace(/<br\s*\/?>/gi, "\n");
            processedText = processedText.replace(/[ ]{2,}$/mg, "\n");
            
            // Apply text without additional formatting
            textFrame.contents = processedText;
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
            chapter: null,
            after: [],
            cover: null
        };
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
            
            if (!this.chapterCount || isNaN(parseInt(this.chapterCount)) || parseInt(this.chapterCount) <= 0) {
                return { valid: false, message: I18n.__('chapterCountPositive') };
            }
            
            if (!this.templates.chapter) {
                return { valid: false, message: I18n.__('chapterTemplateRequired') };
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
        this.generate = function(folder) {
            if (!folder || !folder.exists) {
                throw new Error(I18n.__('destinationFolderNotExist'));
            }
            
            var result = this.validate();
            if (!result.valid) {
                throw new Error(result.message);
            }
            
            try {
                // Create the InDesign book file
                var bookFileName = this.name.toUpperCase() + '-' + I18n.__('bookSuffix') + '.indb';
                var bookFile = new File(folder.fsName + '/' + bookFileName);
                
                // Check if file already exists
                if (bookFile.exists) {
                    var overwrite = confirm(I18n.__('fileExists', bookFileName));
                    if (!overwrite) return false;
                }
                
                var book = app.books.add(bookFile);
                var prefix = this.name.toUpperCase() + '-';
                
                // Generate documents before chapters
                for (var i = 0; i < this.templates.before.length; i++) {
                    this._generateDocument(
                        folder, 
                        this.templates.before[i], 
                        this.templates.before[i].name.replace(/^.*?-/, prefix), 
                        true, 
                        book
                    );
                }
                
                // Generate chapters
                for (var c = 1; c <= parseInt(this.chapterCount); c++) {
                    this._generateDocument(
                        folder,
                        this.templates.chapter,
                        this.templates.chapter.name.replace(/^.*?-/, prefix).replace('.indd', '_' + c + '.indd'),
                        true,
                        book
                    );
                }
                
                // Generate documents after chapters
                for (var j = 0; j < this.templates.after.length; j++) {
                    this._generateDocument(
                        folder,
                        this.templates.after[j],
                        this.templates.after[j].name.replace(/^.*?-/, prefix),
                        true,
                        book
                    );
                }
                
                // Generate cover
                if (this.templates.cover) {
                    this._generateDocument(
                        folder,
                        this.templates.cover,
                        prefix + 'COVER.indd',
                        false,
                        book
                    );
                }
                
                book.save();
                return book;
            } catch (e) {
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
        this._generateDocument = function(folder, template, newName, includeInBook, book) {
            var destFile = new File(folder.fsName + '/' + newName);
            
            try {
                template.copy(destFile.fsName);
                var doc = app.open(destFile, false);
                
                // Add custom variables
                for (var key in this.info) {
                    if (this.info.hasOwnProperty(key)) {
                        BookUtils.Document.addCustomVariable(doc, key, this.info[key]);
                    }
                }
                
                // Replace text placeholders
                BookUtils.Document.replaceTextPlaceholders(doc, this.info, this.displayOptions);
                
                // Replace EAN13 placeholders
                BookUtils.Document.replaceEAN13Placeholders(doc, this.info.isbnPrint, this.info.isbnEbook);
                
                if (this.markdownOptions.injectMarkdown && this.markdownOptions.hasInputFiles) {
                    try {
                        // 1. First inject Markdown
                        this._injectMarkdownContent(doc);
                        
                        // 2. Process overflow without UI manipulation
                        if (doc.isValid) {
                            PageOverflow.processOverflow(doc);
                        }
                    } catch (e) {
                        LogManager.logError("Error processing document", e);
                    }
                }
                
                // Save and close
                doc.save(destFile);
                doc.close();
                
                // Add to book if needed
                if (includeInBook && book) {
                    book.bookContents.add(destFile);
                }
                
                return true;
            } catch (e) {
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
        this._injectMarkdownContent = function(doc) {
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
                    
                    // Remove front matter delimiters if present
                    yamlString = yamlString.replace(/^---\s*\n/, '').replace(/\n---\s*$/, '');
                    
                    var lines = yamlString.split(/\r\n|\r|\n/);
                    var result = {};
                    var currentKey = null;
                    var inList = false;
                    var currentList = [];
                    
                    for (var i = 0; i < lines.length; i++) {
                        var line = lines[i];
                        
                        // Skip empty lines and comments
                        if (localTrim(line) === "" || localTrim(line).charAt(0) === "#") {
                            continue;
                        }
                        
                        // Handle lists
                        if (inList) {
                            if (line.match(/^\s*- /)) {
                                // New list item
                                var listItem = line.replace(/^\s*- /, "");
                                // Remove quotes if present
                                if (listItem.charAt(0) === '"' && listItem.charAt(listItem.length - 1) === '"') {
                                    listItem = listItem.substring(1, listItem.length - 1);
                                }
                                currentList.push(localTrim(listItem));
                                continue;
                            } else {
                                // End of list
                                result[currentKey] = currentList;
                                inList = false;
                                currentList = [];
                                currentKey = null;
                            }
                        }
                        
                        // Normal processing
                        var pair = line.match(/^([^:]+):\s*(.*?)$/);
                        if (pair) {
                            var key = localTrim(pair[1]);
                            var value = localTrim(pair[2]);
                            
                            // Start of a list?
                            if (value === "" && i+1 < lines.length && lines[i+1].match(/^\s*- /)) {
                                currentKey = key;
                                inList = true;
                                currentList = [];
                                continue;
                            }
                            
                            // Normal value
                            result[key] = value;
                        }
                    }
                    
                    // Finalize any in-progress list
                    if (inList && currentKey) {
                        result[currentKey] = currentList;
                    }
                    
                    return result;
                }
        
                // Read YAML file content
                yamlFile.open("r");
                var yamlContent = yamlFile.read();
                yamlFile.close();
        
                var yamlData = parseYaml(yamlContent);
                var inputFiles = yamlData["input-files"];
                
                // Base directory of YAML file (config folder)
                var configDir = yamlFile.parent;
                
                // Determine project root directory (parent of config folder)
                var projectDir = configDir.parent;
                
                // Define possible text folder names with different conventions
                var textFolderVariants = ["text", "Text", "texte", "Texte", "textes", "Textes", "md", "MD", "markdown", "Markdown"];
                
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
                    // Return silently without error
                    return false;
                }
        
                // 2. Find matching Markdown file
                var docName = doc.name;
                var mdFileName = this._findMatchingMarkdownFile(docName, inputFiles);
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
                // CRITICAL: Use \r instead of \n for InDesign line breaks
                targetFrame.contents = mdContent.replace(/\n/g, "\r");
                
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
         * Finds matching Markdown file based on document name
         * @param {string} docName - InDesign document name
         * @param {Array} mdFiles - List of Markdown filenames
         * @return {string|null} Matching Markdown filename or null
         * @private
         */
        this._findMatchingMarkdownFile = function(docName, mdFiles) {
            // Extract descriptive part of InDesign document name
            var descriptivePart = docName.replace(/^.*?-\d+-\d+-/, "")
                                      .replace(/\.indd$/, "")
                                      .toLowerCase()
                                      .replace(/[_-]/g, "");
            
            var bestMatch = null;
            var bestScore = 0;
            
            for (var i = 0; i < mdFiles.length; i++) {
                var mdFile = mdFiles[i];
                
                // Extract descriptive part of Markdown file
                var mdDescriptive = mdFile.replace(/^\d+-/, "")
                                        .replace(/\.md$/, "")
                                        .toLowerCase()
                                        .replace(/[_-]/g, "");
                
                // Calculate match score
                var score = 0;
                
                // Direct match between descriptive parts
                if (descriptivePart === mdDescriptive) {
                    score += 50;  // High score for exact match
                }
                // Partial match 
                else if (descriptivePart.indexOf(mdDescriptive) !== -1 || 
                         mdDescriptive.indexOf(descriptivePart) !== -1) {
                    score += 30;
                }
                
                // Check specific keywords
              var keywords = ["chapter", "intro", "introduction", "conclusion", "appendix", "preface", "postface", "foreword", "index", "bibliography", "glossary", "acknowledgments", "afterword", "epilogue", "prologue", "chapitre", "annexe", "préface", "avant-propos", "bibliographie", "glossaire", "remerciements", "épilogue"];
                
                for (var k = 0; k < keywords.length; k++) {
                    var keyword = keywords[k];
                    if (descriptivePart.indexOf(keyword) !== -1 && mdDescriptive.indexOf(keyword) !== -1) {
                        score += 20;
                        
                        // If both contain a number after keyword, check if it matches
                        var docNumMatch = descriptivePart.match(new RegExp(keyword + "\\s*?(\\d+)", "i"));
                        var mdNumMatch = mdDescriptive.match(new RegExp(keyword + "\\s*?(\\d+)", "i"));
                        
                        if (docNumMatch && mdNumMatch && docNumMatch[1] === mdNumMatch[1]) {
                            score += 30;
                        }
                    }
                }
                
                // Update if best score
                if (score > bestScore) {
                    bestMatch = mdFile;
                    bestScore = score;
                }
            }
            
            return bestScore > 0 ? bestMatch : null;
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
                    "<<Critical_Apparatus>>": bookInfo.critical || "",
                    "<<Translation>>": bookInfo.translation || "",
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
                    "<<Critical_Apparatus>>", 
                    "<<Translation>>", 
                    "<<Original_Title>>", 
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
                    
                    // Mapping between Pandoc fields and BookCreator fields
                    var fieldMappings = {
                        // Standard Pandoc fields
                        "title": "title",
                        "subtitle": "subtitle", 
                        "author": "author",
                        "date": "printDate",
                        "lang": "language",
                        "language": "language",
                        "rights": "rights",
                        "copyright": "rights", // Alias for rights
                        
                        // ISBN specific fields
                        "isbn-print": "isbnPrint",
                        "isbn-ebook": "isbnEbook",
                        "isbnprint": "isbnPrint", // No-hyphen alias
                        "isbnebook": "isbnEbook", // No-hyphen alias
                        
                        // BookCreator specific fields
                        "critical": "critical",
                        "translation": "translation", 
                        "originalTitle": "originalTitle",
                        "coverCredit": "coverCredit",
                        "editions": "editions",
                        "funding": "funding",
                        "price": "price"
                    };
                    
                    var result = {};
                    
                    // Apply mappings
                    for (var srcField in fieldMappings) {
                        if (yamlData.hasOwnProperty(srcField)) {
                            var destField = fieldMappings[srcField];
                            // Ensure there are no leading spaces
                            var value = yamlData[srcField];
                            if (typeof value === "string") {
                                value = value.replace(/^\s+/, "");
                            }
                            result[destField] = value;
                        }
                    }
                    
                    // Include all other fields without specific mapping
                    for (var field in yamlData) {
                        if (!fieldMappings.hasOwnProperty(field)) {
                            // Ensure there are no leading spaces
                            var value = yamlData[field];
                            if (typeof value === "string") {
                                value = value.replace(/^\s+/, "");
                            }
                            result[field] = value;
                        }
                    }
                    
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
                        "printDate": "date",
                        "language": "lang",
                        "rights": "rights",
                        
                        // ISBN specific fields
                        "isbnPrint": "isbn-print",
                        "isbnEbook": "isbn-ebook",
                        
                        // BookCreator specific fields preserved as-is
                        "critical": "critical",
                        "translation": "translation",
                        "originalTitle": "originalTitle", 
                        "coverCredit": "coverCredit",
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
                
                if (yamlData && yamlData["input-files"] && isArray(yamlData["input-files"]) && yamlData["input-files"].length > 0) {
                    result.hasInputFiles = true;
                }
                
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
            
            // Main fields
            win.add('statictext', undefined, I18n.__('fileNamePrefix'));
            var bookNameInput = win.add('edittext', undefined, '');
            
            win.add('statictext', undefined, I18n.__('chapterCount'));
            var chapterCountInput = win.add('edittext', undefined, '');
            
            // Templates panel
            var templatesPanel = win.add('panel', undefined, I18n.__('templates'));
            templatesPanel.alignChildren = 'fill';
            
            // Chapter template
            var chapterTemplateBtn = templatesPanel.add('button', undefined, I18n.__('chapterTemplate'));
            var chapterTemplateText = templatesPanel.add('statictext', undefined, I18n.__('noFileSelected'));
            chapterTemplateBtn.onClick = function() {
                var template = File.openDialog(I18n.__('chapterTemplate') + ' (*.indd)', '*.indd');
                if (template) {
                    book.templates.chapter = template;
                    chapterTemplateText.text = template.name;
                }
            };
            
            // Templates before chapters
            var beforeTemplateBtn = templatesPanel.add('button', undefined, I18n.__('beforeTemplate'));
            var beforeTemplateText = templatesPanel.add('statictext', undefined, I18n.__('filesSelected', 0));
            beforeTemplateBtn.onClick = function() {
                var templates = File.openDialog(I18n.__('beforeTemplate') + ' (*.indd)', '*.indd', true) || [];
                book.templates.before = templates;
                beforeTemplateText.text = I18n.__('filesSelected', templates.length);
            };
            
            // Templates after chapters
            var afterTemplateBtn = templatesPanel.add('button', undefined, I18n.__('afterTemplate'));
            var afterTemplateText = templatesPanel.add('statictext', undefined, I18n.__('filesSelected', 0));
            afterTemplateBtn.onClick = function() {
                var templates = File.openDialog(I18n.__('afterTemplate') + ' (*.indd)', '*.indd', true) || [];
                book.templates.after = templates;
                afterTemplateText.text = I18n.__('filesSelected', templates.length);
            };
            
            // Cover template
            var coverTemplateBtn = templatesPanel.add('button', undefined, I18n.__('coverTemplate'));
            var coverTemplateText = templatesPanel.add('statictext', undefined, I18n.__('noFileSelected'));
            coverTemplateBtn.onClick = function() {
                var template = File.openDialog(I18n.__('coverTemplate') + ' (*.indd)', '*.indd');
                if (template) {
                    book.templates.cover = template;
                    coverTemplateText.text = template.name;
                }
            };
            
            // Book Information button
            var infoBtn = win.add('button', undefined, I18n.__('bookInformation'));
            infoBtn.onClick = function() {
                UI.createInfoWindow(book.info, book.displayOptions, book.markdownOptions, updateMdOptions);
            };
            
            // Markdown section with single checkbox
            var markdownGroup = win.add('group');
            markdownGroup.orientation = "row";
            markdownGroup.alignment = "left";
            markdownGroup.margins = [0, 10, 0, 0];
            
            // Markdown injection checkbox
            var injectMdCheck = markdownGroup.add('checkbox', undefined, I18n.__('injectMarkdown'));
            injectMdCheck.enabled = false;
            
            // Update Markdown options when changed
            injectMdCheck.onClick = function() {
                book.markdownOptions.injectMarkdown = injectMdCheck.value;
            };
            
            // Function to update Markdown options after info window returns
            function updateMdOptions() {
                // Update checkbox state based on detected data
                injectMdCheck.enabled = book.markdownOptions.hasInputFiles;
                
                if (!injectMdCheck.enabled) injectMdCheck.value = false;
            }
            
            // Action buttons
            var actionBtns = win.add('group');
            actionBtns.alignment = 'right';
            actionBtns.add('button', undefined, I18n.__('cancel'), {name:'cancel'});
            var createBtn = actionBtns.add('button', undefined, I18n.__('createBook'), {name:'ok'});
            
            createBtn.onClick = function() {
                // Update book info
                book.name = bookNameInput.text;
                book.chapterCount = chapterCountInput.text;
                
                // Update Markdown options
                book.markdownOptions.injectMarkdown = injectMdCheck.value;
                
                // Validate data
                var validation = book.validate();
                if (!validation.valid) {
                    alert(validation.message);
                    return;
                }
                
                // Generate book
                var folder = Folder.selectDialog(I18n.__('chooseDestionationFolder'));
                if (folder) {
                    try {
                        // Disable InDesign native error messages during execution
                        var originalUserInteractionLevel = app.scriptPreferences.userInteractionLevel;
                        app.scriptPreferences.userInteractionLevel = UserInteractionLevels.NEVER_INTERACT;
                        
                        var result = book.generate(folder);
                        
                        // Restore user interaction level BEFORE opening the book
                        app.scriptPreferences.userInteractionLevel = originalUserInteractionLevel;
                        
                        if (result) {
                            alert(I18n.__('bookGenerated'));
                            win.close();
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
        createInfoWindow: function(bookInfo, displayOptions, markdownOptions, updateMdDiagnosticCallback) {
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
            
            // Simplified Markdown diagnostics
            var mdDiagnosticGroup = col1.add("group");
            mdDiagnosticGroup.orientation = "column";
            mdDiagnosticGroup.alignChildren = "left";
            mdDiagnosticGroup.margins = [0, 10, 0, 0];
            
            var mdDiagnosticLabel = mdDiagnosticGroup.add("statictext", undefined, I18n.__('markdownDetection'));
            var mdDiagnostic = mdDiagnosticGroup.add("statictext", undefined, 
                markdownOptions.yamlPath ? 
                (markdownOptions.hasInputFiles ? I18n.__('inputFilesDetected') : I18n.__('inputFilesNotDetected')) : 
                I18n.__('noYAMLLoaded'));
            mdDiagnostic.preferredSize.width = 250;
            
            // Column 2
            var col2 = inputGroup.add("group");
            col2.orientation = "column";
            col2.alignChildren = "left";
            col2.spacing = 5;
            
            col2.add("statictext", undefined, I18n.__('criticalApparatus'));
            var criticalInput = col2.add("edittext", undefined, bookInfo.critical || "", {multiline: true});
            criticalInput.preferredSize = [250, 40];
            
            col2.add("statictext", undefined, I18n.__('translation'));
            var translationInput = col2.add("edittext", undefined, bookInfo.translation || "", {multiline: true});
            translationInput.preferredSize = [250, 40];
            
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
                        
                        // Store YAML data and check Markdown elements
                        markdownOptions.yamlPath = importResult.yamlPath;
                        markdownOptions.yamlMeta = importResult.yamlMeta;
                        
                        var mdElements = BookUtils.File.detectMarkdownElements(importResult.yamlMeta);
                        markdownOptions.hasInputFiles = mdElements.hasInputFiles;
                        
                        // Update diagnostic text in this window
                        mdDiagnostic.text = markdownOptions.hasInputFiles ? 
                                           I18n.__('inputFilesDetected') : 
                                           I18n.__('inputFilesNotDetected');
                                            
                        // Update Markdown options in main window
                        if (updateMdDiagnosticCallback) {
                            updateMdDiagnosticCallback();
                        }
                        
                        alert(I18n.__('yamlImportCompleted'));
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
