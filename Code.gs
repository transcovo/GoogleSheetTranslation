//
// Code.gs
//
// Copyright © 2015 Sebastien MICHOY and contributors.
//
// Redistribution and use in source and binary forms, with or without
// modification, are permitted provided that the following conditions are met:
//
// Redistributions of source code must retain the above copyright notice, this
// list of conditions and the following disclaimer. Redistributions in binary
// form must reproduce the above copyright notice, this list of conditions and
// the following disclaimer in the documentation and/or other materials
// provided with the distribution. Neither the name of the nor the names of
// its contributors may be used to endorse or promote products derived from
// this software without specific prior written permission.
//
// THIS SOFTWARE IS PROVIDED BY THE COPYRIGHT HOLDERS AND CONTRIBUTORS "AS IS"
// AND ANY EXPRESS OR IMPLIED WARRANTIES, INCLUDING, BUT NOT LIMITED TO, THE
// IMPLIED WARRANTIES OF MERCHANTABILITY AND FITNESS FOR A PARTICULAR PURPOSE
// ARE DISCLAIMED. IN NO EVENT SHALL THE COPYRIGHT HOLDER OR CONTRIBUTORS BE
// LIABLE FOR ANY DIRECT, INDIRECT, INCIDENTAL, SPECIAL, EXEMPLARY, OR
// CONSEQUENTIAL DAMAGES (INCLUDING, BUT NOT LIMITED TO, PROCUREMENT OF
// SUBSTITUTE GOODS OR SERVICES; LOSS OF USE, DATA, OR PROFITS; OR BUSINESS
// INTERRUPTION) HOWEVER CAUSED AND ON ANY THEORY OF LIABILITY, WHETHER IN
// CONTRACT, STRICT LIABILITY, OR TORT (INCLUDING NEGLIGENCE OR OTHERWISE)
// ARISING IN ANY WAY OUT OF THE USE OF THIS SOFTWARE, EVEN IF ADVISED OF THE
// POSSIBILITY OF SUCH DAMAGE.

/** System Functions **/

/**
 * A special function that runs when the spreadsheet is open, used to add a
 * custom menu to the spreadsheet.
 */
function onOpen() {
  var spreadsheet = SpreadsheetApp.getActive();
  var menuItems = [{name: 'Generate translation', functionName: 'displayGenerationUI'},
                   {name: 'Initialize spreadsheet', functionName: 'displayInitiationSpreadsheetUI'},
                  {name: 'Validate Spreadsheet', functionName: 'validateSpreadsheet'}];

  spreadsheet.addMenu('Translation', menuItems);
}

/** UI Functions **/

/**
 * Displays the UI to generate translation files.
 */
function displayGenerationUI() {
  var previewTitle = "Preview"

  var html = HtmlService.createHtmlOutputFromFile('Preview').setSandboxMode(HtmlService.SandboxMode.IFRAME);
  SpreadsheetApp.getUi().showModalDialog(html, previewTitle);
}

/**
 * Displays the UI to initialize new sheets.
 */
function displayInitiationSpreadsheetUI() {
  var ui = SpreadsheetApp.getUi()
  var response = ui.alert("Do you want initialize the spreadsheet?\nWARNING: It can erase some of your data.", ui.ButtonSet.YES_NO);

  if (response == ui.Button.YES) {
    generateTranslationSpreadsheet();
  }
}

/**
 * Validate that cells containing a format / variable are correct
 */
function validateSpreadsheet() {
  var validationColumn = 4; // column containing format definition
  var errorValidation = 0;
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getDataRange().getValues();

  for (var i = 0; i < data.length; i++) {
    formatDesc = data[i][validationColumn];

    if(formatDesc != "") {
      Logger.log('validateSpreadsheet.Format: ' + formatDesc);
      errorValidation += validateFormat(formatDesc,i+1);
    }
  }

  var ui = SpreadsheetApp.getUi()

  if(errorValidation == 0) {
    ui.alert("Congrats, your file is ready to build", ui.ButtonSet.OK);
  } else {
    ui.alert("You have "+errorValidation+" error(s) on your document. Search for red background cells ;)", ui.ButtonSet.OK);
  }
}

/**
 * for one line, check that cells variable format is ok
 */
function validateFormat(formatDesc,lineIndex) {
  var startingColumn = 2;
  var error = 0;
  var sheet = SpreadsheetApp.getActiveSheet();
  var data = sheet.getRange(lineIndex, startingColumn, 1, 2).getValues();
  var argsNeeded = formatDesc.split("|"); // split format definition column

  for(var col = 0; col < data[0].length; col++) {
    // the regexp search for :
    // %1$d or %2$d or %1$s => valid ony for $s or $d
    // or
    // #?1$d# or #?2$d# or #?1$s# => valid ony for $s or $d
    // or
    // #?d# or #?#
    matches = data[0][col].match(/%(\d)\$[d,s]|#\?(\d)\$[d,s]#|#\?[d,s]?#/g);

    if(JSON.stringify(argsNeeded).localeCompare(JSON.stringify(matches)) == 0) {
      sheet.getRange(lineIndex, startingColumn+col, 1, 2).setBackground("white");
    } else {
      error++;
      sheet.getRange(lineIndex, startingColumn+col, 1, 1).setBackground("red");
    }
  }

  return error;
}

/** Translation Functions **/

/**
 * Generates the translation files.
 *
 * Return them as following:
 *
 * [
 *   {
 *     osName: "Android",
 *     translations: [
 *       {
 *         language: "English",
 *         fileName: "strings.xml"
 *         file: "The translated text"
 *       }
 *     ]
 *   }
 * ]
 *
 */
function generateTranslationFiles() {
  var oss = getOSs();
  var languages = getLanguages();
  var translatedFiles = [];

  for (var osIndex = 0, ossLength = oss.length; osIndex < ossLength; ++osIndex) {
    var os = oss[osIndex];
    var translatedFilesPerOS = [];

    for (var languageIndex = 0, languagesLength = languages.length; languageIndex < languagesLength; ++languageIndex) {
      var language = languages[languageIndex];

      var translatedFileData = getTranslationData(languageIndex + 2, os);
      var translatedFile = generateTranslationFile(os, language['name'], translatedFileData);

      translatedFilesPerOS.push({language: language['name'], fileName: getFileName(os, language['name']), file: translatedFile});
    }

    translatedFiles.push({osName: os, translations: translatedFilesPerOS});
  }

  return translatedFiles;
}

/**
 * Gets data for the translation
 *
 * Return them as following:
 *
 * [
 *   {
 *     sectionTitle: "My section title",
 *     values: [
 *       {
 *         key: "translationKey",
 *         value: "My translation"
 *       }
 *     ]
 *   }
 * ]
 */
function getTranslationData(languageColumnNumber, os) {
  var sheet = getTranslationSheet();

  if (sheet == null)
    return "";

  var firstTranslationRowNumber = 4;
  var keysColumnNumber = 1;
  var osColumnNumber = getLanguages().length + 2;
  var numberOfRows = sheet.getMaxRows() - firstTranslationRowNumber + 1;

  var keysColumn = sheet.getRange(firstTranslationRowNumber, keysColumnNumber, numberOfRows, 1).getValues();
  var translationsColumn = sheet.getRange(firstTranslationRowNumber, languageColumnNumber, numberOfRows, 1).getValues();
  var osColumn = sheet.getRange(firstTranslationRowNumber, osColumnNumber, numberOfRows, 1).getValues();

  var translations = []

  for (var i = 0; i < keysColumn.length; ++i) {
    if (osColumn[i][0].length == 0) { /* No OS means it is a section. */
      translations.push({ 'sectionTitle': keysColumn[i][0] });
    } else if (matchForOS(os, osColumn[i][0])) { /* The line is useful for the os, so we add it. */
      var translation = { 'key': keysColumn[i][0], 'value': translationsColumn[i][0] };
      var section = {}

      if (translations.length > 0) {
        section = translations[translations.length - 1]
      }

      if (section['values'] == undefined) {
        section['values'] = []
      }

      section['values'].push(translation);
    }
  }

  var translations = translations.filter(function (section) { return (section['values'] != undefined); });

  return translations
}

/** Generation Translation Files Functions **/

/**
 * Generates the Android translation files.
 */
function generateAndroidTranslationFile(os, language, data) {
  var translationContent = '<!---\n' + getHeader(os, language) + '\n-->\n\n<resources>\n';
  var pluralsKey = null;

  function checkPlurals(isPlural, key) {
    if (isPlural) {
      if (pluralsKey == null) {
        translationContent += '<plurals name="' + key + '">\n';
      } else if (pluralsKey !== key) {
        translationContent += '</plurals>\n<plurals name="' + key + '">\n';
      }
      pluralsKey = key
    } else {
      if (pluralsKey != null) {
        translationContent += '</plurals>\n';
        pluralsKey = null;
      }
    }
  }

  function isPlural(key){
    return key.indexOf("_plurals_")>-1;
  }


  for (var sectionNumber = 0; sectionNumber < data.length; sectionNumber++) {
    var section = data[sectionNumber];

    if (section['sectionTitle'] != undefined){
      checkPlurals(false);
      translationContent += '\n<!-- ' + section['sectionTitle'] + ' -->\n\n';
    }

    if (section['values'] != undefined) {
      for (var translationNumber = 0; translationNumber < section['values'].length; translationNumber++) {
        var translation = section['values'][translationNumber];

        if (translation != undefined && translation['key'] != undefined && translation['value'] != undefined) {
          if (isPlural(translation['key'])) {
            var keyParts = translation['key'].split("_plurals_");
            var key = keyParts[0];
            var quantity = keyParts[1];
            var value = translation['value'].replace(/#(\d*?)\?(.*?)#/g, typeReplacerAndroid).replace(/'/g, '\\\'').replace(/"/g, '\\\"')
            checkPlurals(true, key)
            translationContent += '<item quantity="' + quantity + '">' + value + '</item>\n'
          } else {
            var key = translation['key']
            var value = translation['value'].replace(/#(\d*?)\?(.*?)#/g, typeReplacerAndroid).replace(/'/g, '\\\'').replace(/"/g, '\\\"')
            checkPlurals(false)
            translationContent += '<string name="' + key + '">' + value + '</string>\n'
          }
        }
      }
    }
  }

  translationContent += '</resources>'

  return translationContent;
}

function typeReplacerAndroid(match, p1, p2, offset, string) {
  var stringFormatted = '%';
  if (p1)
    stringFormatted += p1 + '$';
  if (p2)
    stringFormatted += p2;
  else
    stringFormatted += 's';
  return stringFormatted;
}

function typeReplacerIOS(match, p1, p2, offset, string) {
  var stringFormatted = '%';
  if (p1)
    stringFormatted += p1 + '$';
  if (p2)
    stringFormatted += p2;
  else
    stringFormatted += '@';
  return stringFormatted;
}

/**
 * Generates the iOS translation files.
 */
function generateIOSTranslationFile(os, language, data) {
  var translationContent = '';

  translationContent = '/*\n' + getHeader(os, language) + '\n*/\n'

  for (var sectionNumber = 0; sectionNumber < data.length; sectionNumber++) {
    var section = data[sectionNumber];

    if (section['sectionTitle'] != undefined)
      translationContent += '\n/* ' + section['sectionTitle'] + ' */\n\n';

    if (section['values'] != undefined) {
      for (var translationNumber = 0; translationNumber < section['values'].length; translationNumber++) {
        var translation = section['values'][translationNumber];

        if (translation != undefined && translation['key'] != undefined && translation['value'] != undefined) {
          var key = translation['key'].split("_").join(".")
          var value = translation['value'].replace(/#(\d*?)\?(.*?)#/g, typeReplacerIOS).replace(/"/g, '\\\"')
          translationContent += '"' + key + '" = "' + value + '";\n';
        }
      }
    }
  }

  return translationContent;
}

/**
 * Calls the right generation function for depending of os and language.
 */
function generateTranslationFile(os, language, data) {
  var translatedFile = ''

  switch (os) {
    case 'Android':
      translatedFile = generateAndroidTranslationFile(os, language, data);
      break;
    case 'iOS':
      translatedFile = generateIOSTranslationFile(os, language, data);
      break;
  }

  return translatedFile;
}

/** Generation Sheets Functions **/

/**
 * Generate the translation sheet.
 */
function generateTranslationSpreadsheet() {
  var spreadsheet = SpreadsheetApp.getActive();
  var sheet = getTranslationSheet();

  if (sheet == null) {
    sheet = spreadsheet.insertSheet(getTranslationSheetName());
  }

  var languages = getLanguages();

  var numberOfColumns = 2 + languages.length;
  var numberOfRows = 5;

  /* Clear the sheet */
  sheet.clear()

  /* Resize the sheet */
  if (sheet.getMaxColumns() > 1) {
    sheet.deleteColumns(1, sheet.getMaxColumns() - 1);
  }

  if (sheet.getMaxRows() > 1) {
    sheet.deleteRows(1, sheet.getMaxRows() - 1);
  }

  sheet.getRange(1, 1, 1, 1).clearDataValidations();
  sheet.insertColumns(1, numberOfColumns - 1);
  sheet.insertRows(1, numberOfRows - 1);

  /* Resize cells */
  sheet.setColumnWidth(1, 300);
  sheet.setColumnWidth(sheet.getMaxColumns(), 100);

  for (var i = 0, languageColumnsNumber = languages.length; i < languageColumnsNumber; ++i) {
    sheet.setColumnWidth(i + 2, 450);
  }

  for (var i = 0, rowsNumber = sheet.getMaxRows(); i < rowsNumber; ++i) {
    sheet.setRowHeight(i + 1, 30);
  }

  /* General Fonts */
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).setFontFamily('Trebuchet MS');
  sheet.getRange(1, 1, sheet.getMaxRows(), sheet.getMaxColumns()).setFontSize(10);

  /* Title Cell */
  sheet.setRowHeight(1, 50)
  sheet.getRange(1, 1, 1, sheet.getMaxColumns()).merge();
  sheet.getRange(1, 1, 1, sheet.getMaxColumns()).setFontSize(18);
  sheet.getRange(1, 1, 1, sheet.getMaxColumns()).setFontWeight('bold');
  sheet.getRange(1, 1, 1, sheet.getMaxColumns()).setBackground('#E59142');
  sheet.getRange(1, 1, 1, sheet.getMaxColumns()).setFontColor('white');
  sheet.getRange(1, 1, 1, sheet.getMaxColumns()).setVerticalAlignment('middle');
  sheet.getRange(1, 1, 1, 1).setValue(getAppName());

  /* Headers Cells */
  sheet.setRowHeight(3, 30)
  sheet.getRange(3, 1, 1, sheet.getMaxColumns()).setFontWeight('bold');
  sheet.getRange(3, 1, 1, sheet.getMaxColumns()).setHorizontalAlignment('center');
  sheet.getRange(3, 1, 1, sheet.getMaxColumns()).setBackground('#E59142');
  sheet.getRange(3, 1, 1, sheet.getMaxColumns()).setFontColor('white');
  sheet.getRange(3, 1, 1, sheet.getMaxColumns()).setVerticalAlignment('middle');

  sheet.getRange(3, 1, 1, 1).setValue('Keys');
  sheet.getRange(3, sheet.getMaxColumns(), 1, 1).setValue('OS');

  for (var i = 0, languageColumnsNumber = languages.length; i < languageColumnsNumber; ++i) {
    sheet.getRange(3, i + 2, 1, 1).setValue(languages[i]['displayName']);
  }

  /* Section Cell */
  sheet.setRowHeight(4, 30)
  sheet.getRange(4, 1, 1, sheet.getMaxColumns()).merge();
  sheet.getRange(4, 1, 1, sheet.getMaxColumns()).setBackground('#F5B271');
  sheet.getRange(4, 1, 1, sheet.getMaxColumns()).setFontColor('white');
  sheet.getRange(4, 1, 1, sheet.getMaxColumns()).setVerticalAlignment('middle');
  sheet.getRange(4, 1, 1, 1).setValue('Launch Screen');

  /* Traduction Cells */
  var ossList = getOSs();
  ossList.unshift('All');

  var ossValidation = SpreadsheetApp.newDataValidation().requireValueInList(ossList, true).setAllowInvalid(false).build();

  sheet.setRowHeight(5, 30)
  sheet.getRange(5, 1, 1, sheet.getMaxColumns()).setVerticalAlignment('top');

  sheet.getRange(5, 1, 1, 1).setFontFamily('Courier New');
  sheet.getRange(5, 1, 1, 1).setFontColor('#DF6769');
  sheet.getRange(5, 1, 1, 1).setValue('welcome_text');

  sheet.getRange(5, sheet.getMaxColumns(), 1, 1).setHorizontalAlignment('center');
  sheet.getRange(5, sheet.getMaxColumns(), 1, 1).setValue('All');
  sheet.getRange(5, sheet.getMaxColumns(), 1, 1).setDataValidation(ossValidation);

  for (var i = 0, languageColumnsNumber = languages.length; i < languageColumnsNumber; ++i) {
    sheet.getRange(5, i + 2, 1, 1).setValue('Hello World!');
  }
}

/** Tools Functions **/

/**
 * Returns the application name.
 */
function getAppName() {
  return "MyLocalizedApp"
}

/**
 * Returns the file name in function of the OS and the language.
 */
function getFileName(os, language) {
  var fileName = 'translation'

  switch (os) {
    case 'Android':
        fileName = 'string.xml';
        break;
    case 'iOS':
      fileName = 'Localizable.strings';
      break;
  }

  return fileName
}

/**
 * Returns the list of languages.
 */
function getLanguages() {
  return [{name: 'English', displayName: 'English (GB)'},
          {name: 'French', displayName: 'French (FR)'}];
}

/**
 * Returns the list of OSs.
 */
function getOSs() {
  return ['Android', 'iOS'];
}

/**
 * Returns the header for generated files.
 */
function getHeader(os, language) {
  var date = new Date();
  var currentDate = (date.getMonth() + 1) + '/' + date.getDate() + '/' + date.getFullYear()

  var header = ''
  header += '  ' + getFileName(os, language) + '\n';
  header += '  ' + getAppName() + '\n\n';
  header += '  Generated by AppTranslation on ' + currentDate + '\n';
  header += '  Copyright © 2015 Sébastien MICHOY and contributors.';

  return header
}

function getTranslationSheet() {
  var spreadsheet = SpreadsheetApp.getActive();

  return spreadsheet.getSheetByName(getTranslationSheetName());
}

/**
 * Return the name of the translation sheet.
 */
function getTranslationSheetName() {
  return "Translations";
}

/**
 * Indicates if the `os` and the `osString` are compliant.
 *
 * Return `true` if they are compliant, else, returns `false`.
 */
function matchForOS(os, osString) {
  if (os.length == 0 || osString.length == 0)
    return false;

  if (osString == 'All')
    return true;

  if (osString == os)
    return true;

  return false;
}
