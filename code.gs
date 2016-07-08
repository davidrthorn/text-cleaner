// Create Text Cleaner sub-menu menu in the add-on menu
function onInstall(e) {
  onOpen(e);
}

function onOpen(e) {
    DocumentApp.getUi()
      .createAddonMenu()
      .addItem('Clean selected text', 'cleanText')
      .addItem('Configure', 'showDialog')
      .addSeparator()
      .addItem('Remove links and underlining', 'removeLinks')
      .addItem('Remove line and paragraph breaks', 'removeBoth')
      .addItem('Remove multiple spaces', 'removeMultipleSpaces')
      .addItem('Remove tabs', 'removeTabs')
      .addToUi();
  }
  // Open options dialog from add-on menu

function showDialog() {
    var html = HtmlService.createHtmlOutputFromFile('dialog')
      .setSandboxMode(HtmlService.SandboxMode.IFRAME)
      .setWidth(360)
      .setHeight(240);
    DocumentApp.getUi() // Or DocumentApp or FormApp.
      .showModalDialog(html, 'Configure Text Cleaner');
  }
  // Interact with saved style settings

function getUserSettings() {
    var userProperties = PropertiesService.getUserProperties();
    var selection = DocumentApp.getActiveDocument()
      .getSelection();
    if (selection) {
      var selected = 'true';
    } else {
      var selected = 'false';
    }
    var user_settings = {
      bold: userProperties.getProperty('CLEAR_bold'),
      italic: userProperties.getProperty('CLEAR_italic'),
      underline: userProperties.getProperty('CLEAR_underline'),
      strike: userProperties.getProperty('CLEAR_strike'),
      indent: userProperties.getProperty('CLEAR_indent'),
      links: userProperties.getProperty('CLEAR_links'),
      line_breaks: userProperties.getProperty('CLEAR_line_breaks'),
      paras: userProperties.getProperty('CLEAR_paras'),
      multiple: userProperties.getProperty('CLEAR_multiple'),
      tabs: userProperties.getProperty('CLEAR_tabs'),
      selected: selected
    };
    return user_settings;
  }
  // Update document from the options dialog (first sets new document properties)

function updateDocument(bold, italic, underline, strike, indent, links, line_breaks, paras, multiple, tabs) {
  var userProperties = PropertiesService.getUserProperties();
  userProperties.setProperty('CLEAR_bold', bold);
  userProperties.setProperty('CLEAR_italic', italic);
  userProperties.setProperty('CLEAR_underline', underline);
  userProperties.setProperty('CLEAR_strike', strike);
  userProperties.setProperty('CLEAR_indent', indent);
  userProperties.setProperty('CLEAR_links', links);
  userProperties.setProperty('CLEAR_line_breaks', line_breaks);
  userProperties.setProperty('CLEAR_paras', paras);
  userProperties.setProperty('CLEAR_multiple', multiple);
  userProperties.setProperty('CLEAR_tabs', tabs);
}

function updateClear(bold, italic, underline, strike, indent, links, line_breaks, paras, multiple, tabs) {
    var userProperties = PropertiesService.getUserProperties();
    userProperties.setProperty('CLEAR_bold', bold);
    userProperties.setProperty('CLEAR_italic', italic);
    userProperties.setProperty('CLEAR_underline', underline);
    userProperties.setProperty('CLEAR_strike', strike);
    userProperties.setProperty('CLEAR_indent', indent);
    userProperties.setProperty('CLEAR_links', links);
    userProperties.setProperty('CLEAR_line_breaks', line_breaks);
    userProperties.setProperty('CLEAR_paras', paras);
    userProperties.setProperty('CLEAR_multiple', multiple);
    userProperties.setProperty('CLEAR_tabs', tabs);
    cleanText();
  }
  // Update document from the add-on menu

function cleanText() {
  var userProperties = PropertiesService.getUserProperties();
  var CLEAR_bold = userProperties.getProperty('CLEAR_bold');
  var CLEAR_italic = userProperties.getProperty('CLEAR_italic');
  var CLEAR_underline = userProperties.getProperty('CLEAR_underline');
  var CLEAR_strike = userProperties.getProperty('CLEAR_strike');
  var CLEAR_indent = userProperties.getProperty('CLEAR_indent');
  var CLEAR_links = userProperties.getProperty('CLEAR_links');
  var CLEAR_line_breaks = userProperties.getProperty('CLEAR_line_breaks');
  var CLEAR_paras = userProperties.getProperty('CLEAR_paras');
  var CLEAR_multiple = userProperties.getProperty('CLEAR_multiple');
  var CLEAR_tabs = userProperties.getProperty('CLEAR_tabs');
  var selection = DocumentApp.getActiveDocument()
    .getSelection();
  if (selection) {
    var elements = selection.getRangeElements();
    for (var i = 0; i < elements.length; i++) {
      var element = elements[i];
      // Only deal with text elements
      if (element.getElement()
        .editAsText) {
        var text = element.getElement()
          .editAsText();
        var style = {};
        // Add user's unwanted attributes to style array
        if (CLEAR_bold != 'true') {
          style[DocumentApp.Attribute.BOLD] = null;
        }
        if (CLEAR_italic != 'true') {
          style[DocumentApp.Attribute.ITALIC] = null;
        }
        if (CLEAR_underline != 'true') {
          style[DocumentApp.Attribute.UNDERLINE] = null;
        }
        if (CLEAR_strike != 'true') {
          style[DocumentApp.Attribute.STRIKETHROUGH] = null;
        }
        if (CLEAR_indent != 'true') {
          style[DocumentApp.Attribute.INDENT_END] = null;
        }
        if (CLEAR_links == 'true') {
          style[DocumentApp.Attribute.LINK_URL] = null;
        }
        // Add all standard clearable attributes to style array
        style[DocumentApp.Attribute.FONT_SIZE] = null;
        style[DocumentApp.Attribute.HORIZONTAL_ALIGNMENT] = null;
        style[DocumentApp.Attribute.LINE_SPACING] = null;
        style[DocumentApp.Attribute.SPACING_BEFORE] = null;
        style[DocumentApp.Attribute.SPACING_AFTER] = null;
        style[DocumentApp.Attribute.FOREGROUND_COLOR] = null;
        style[DocumentApp.Attribute.BACKGROUND_COLOR] = null;
        style[DocumentApp.Attribute.FONT_FAMILY] = null;
        // Deal with partially selected text
        if (element.isPartial()) {
          text.setAttributes(element.getStartOffset(), element.getEndOffsetInclusive(), style);
          if (CLEAR_line_breaks == 'true') {
            var start = element.getStartOffset();
            var finish = element.getEndOffsetInclusive();
            var oldText = text.getText()
              .slice(start, finish);
            if (oldText.match(/\r/)) {
              var number = oldText.match(/\r/g)
                .length;
              for (var j = 0; j < number; j++) {
                var location = oldText.search(/\r/);
                text.deleteText(start + location, start + location);
                text.insertText(start + location, ' ');
                var oldText = oldText.replace(/\r/, ' ');
              }
            }
          }
          if (CLEAR_paras == 'true') {
            var type = element.getElement()
              .getParent()
              .getType();
            if (type == "PARAGRAPH") {
              if (CLEAR_multiple == 'true' && text.getText()
                .length > 0) {
                text.replaceText("[ ][ ]+", " ");
                var firstChar = text.getText()
                  .charAt(0);
                if (firstChar == " ") {
                  text.deleteText(0, 0);
                }
              }
              var para = element.getElement()
                .getParent()
              var paratext = para.asText()
                .getText();
              var paralength = paratext.length;
              var prev = para.getPreviousSibling();
              var finalchar = paratext.charAt(paralength - 1);
              if (paralength > 0 && finalchar == " ") {
                text.deleteText(paralength - 1, paralength - 1);
              }
              var parastyle = para.getAttributes()
                .HEADING;
              if (i == 0 && paralength == 0) {
                para.removeFromParent();
              }
              if (i != elements.length - 1 && paralength == 0 && nextparastyle != "Normal") {
                para.removeFromParent();
              };
              if (parastyle == "Normal") {
                if (prev) {
                  var prevparastyle = prev.getAttributes()
                    .HEADING;
                }
                if (i > 0 && prev.getType() == "PARAGRAPH" && prevparastyle == "Normal") {
                  var check = para.getPreviousSibling()
                    .asText()
                    .getText()
                    .length;
                  if (check > 0) {
                    para.getPreviousSibling()
                      .asParagraph()
                      .appendText(" ");
                  }
                  para.merge();
                }
              }
            }
          }
        }
        // Deal with fully selected text
        else {
          var theelement = element.getElement();
          if (theelement != "ListItem") {
            if (CLEAR_indent != 'true') {
              style[DocumentApp.Attribute.INDENT_START] = null;
              style[DocumentApp.Attribute.INDENT_FIRST_LINE] = null;
            }
          } else {
            var nest = theelement.asListItem()
              .getNestingLevel();
            var newstart = nest * 36 + 36;
            var newfirst = nest * 36 + 18;
            style[DocumentApp.Attribute.INDENT_START] = newstart;
            style[DocumentApp.Attribute.INDENT_FIRST_LINE] = newfirst;
          }
          text.setAttributes(style);
          var type = theelement.getType();
          if (type == "TABLE_CELL") {
            var numchi = element.getElement()
              .asTableCell()
              .getNumChildren();
            for (var p = numchi - 1; p >= 0; p--) {
              var para = theelement.asTableCell()
                .getChild(p);
              var celltext = para.editAsText();
              para.setAttributes(style);
              celltext.setAttributes(style);
            }
          }
          if (type == "TABLE") {
            var numrows = theelement.asTable()
              .getNumChildren();
            for (var q = 0; q < numrows; q++) {
              var row = element.getElement()
                .asTable()
                .getChild(q)
                .asTableRow();
              var numcells = row.getNumChildren();
              for (var r = 0; r < numcells; r++) {
                var cell = row.getChild(r)
                  .asTableCell();
                var numpara = cell.getNumChildren();
                for (var p = numpara - 1; p >= 0; p--) {
                  var para = cell.getChild(p);
                  var celltext = para.editAsText();
                  para.setAttributes(style);
                  celltext.setAttributes(style);
                }
              }
            }
          }
          if (CLEAR_line_breaks == 'true') {
            text.replaceText("\\v+", " ");
          }
          if (CLEAR_paras == 'true') {
            var type = element.getElement()
              .getType();
            if (type == "PARAGRAPH") {
              if (CLEAR_multiple == 'true' && text.getText()
                .length > 0) {
                text.replaceText("[ ][ ]+", " ");
                var firstChar = text.getText()
                  .charAt(0);
                if (firstChar == " ") {
                  text.deleteText(0, 0);
                }
              }
              var para = element.getElement();
              var paratext = para.asText()
                .getText();
              var paralength = paratext.length;
              var prev = para.getPreviousSibling();
              var finalchar = paratext.charAt(paralength - 1);
              if (paralength > 0 && finalchar == " ") {
                text.deleteText(paralength - 1, paralength - 1);
              }
              var parastyle = para.getAttributes()
                .HEADING;
              if (para.getNextSibling()) {
                var nextparastyle = para.getNextSibling()
                  .getAttributes();
              }
              if (i == 0 && paralength == 0 && nextparastyle != "Normal") {
                para.removeFromParent();
              }
              if (parastyle == "Normal") {
                if (prev) {
                  var prevparastyle = prev.getAttributes()
                    .HEADING;
                }
                if (prev && i > 0 && prev.getType() == "PARAGRAPH" && prevparastyle == "Normal") {
                  var prevlength = prev.asText()
                    .getText()
                    .length;
                  if (paralength > 0) {
                    prev.asParagraph()
                      .appendText(" ");
                  }
                  para.merge();
                }
              }
            } else if (type == "TABLE_CELL") {
              var numchi = theelement.asTableCell()
                .getNumChildren();
              for (var p = numchi - 1; p >= 0; p--) {
                var para = theelement.asTableCell()
                  .getChild(p);
                var celltext = para.editAsText();
                var prev = para.getPreviousSibling();
                var paratext = para.asText()
                  .getText();
                var paralength = paratext.length;
                if (prev) {
                  var prevtext = prev.asText()
                    .getText();
                  var prevlength = prevtext.length;
                  var prevchar = prevtext.charAt(prevlength - 1);
                }
                if (prev && prevlength > 0 && prevchar == " ") {
                  if (CLEAR_multiple == 'true') {
                    var newprevtext = prev.asText()
                      .replaceText("[ ][ ]+", " ");
                    var newlength = newprevtext.getText()
                      .length;
                    if (newlength != prevlength) {
                      prev.asText()
                        .deleteText(newlength - 1, newlength - 1);
                    } else {
                      prev.asText()
                        .deleteText(prevlength - 1, prevlength - 1);
                    }
                  } else {
                    prev.asText()
                      .deleteText(prevlength - 1, prevlength - 1);
                  }
                }
                if (CLEAR_multiple == 'true' && celltext.getText()
                  .length > 0) {
                  celltext.replaceText("[ ][ ]+", " ");
                  var firstChar = celltext.getText()
                    .charAt(0);
                  if (firstChar == " ") {
                    celltext.deleteText(0, 0);
                  }
                }
                var parastyle = para.getAttributes()
                  .HEADING;
                if (para.getNextSibling()) {
                  var nextparastyle = para.getNextSibling()
                    .getAttributes();
                }
                if (p == 0 && paralength == 0 && nextparastyle != "Normal") {
                  para.removeFromParent();
                }
                if (parastyle == "Normal") {
                  if (prev) {
                    var prevparastyle = prev.getAttributes()
                      .HEADING;
                  }
                  if (prev && prev.getType() == "PARAGRAPH" && prevparastyle == "Normal") {
                    if (paralength > 0) {
                      para.getPreviousSibling()
                        .asParagraph()
                        .appendText(" ");
                    }
                    para.merge();
                  }
                }
              }
            } else if (type == "TABLE") {
              var numrows = element.getElement()
                .asTable()
                .getNumChildren();
              for (var q = 0; q < numrows; q++) {
                var row = element.getElement()
                  .asTable()
                  .getChild(q)
                  .asTableRow();
                var numcells = row.getNumChildren();
                for (var r = 0; r < numcells; r++) {
                  var cell = row.getChild(r)
                    .asTableCell();
                  var numpara = cell.getNumChildren();
                  for (var p = numpara - 1; p >= 0; p--) {
                    var para = cell.asTableCell()
                      .getChild(p);
                    var celltext = para.editAsText();
                    var prev = para.getPreviousSibling();
                    var paratext = para.asText()
                      .getText();
                    var paralength = paratext.length;
                    if (prev) {
                      var prevtext = prev.asText()
                        .getText();
                      var prevlength = prevtext.length;
                      var prevchar = prevtext.charAt(prevlength - 1);
                    }
                    if (prev && prevlength > 0 && prevchar == " ") {
                      if (CLEAR_multiple == 'true') {
                        var newprevtext = prev.asText()
                          .replaceText("[ ][ ]+", " ");
                        var newlength = newprevtext.getText()
                          .length;
                        if (newlength != prevlength) {
                          prev.asText()
                            .deleteText(newlength - 1, newlength - 1);
                        } else {
                          prev.asText()
                            .deleteText(prevlength - 1, prevlength - 1);
                        }
                      } else {
                        prev.asText()
                          .deleteText(prevlength - 1, prevlength - 1);
                      }
                    }
                    if (CLEAR_multiple == 'true' && celltext.getText()
                      .length > 0) {
                      celltext.replaceText("[ ][ ]+", " ");
                      var firstChar = celltext.getText()
                        .charAt(0);
                      if (firstChar == " ") {
                        celltext.deleteText(0, 0);
                      }
                    }
                    var parastyle = para.getAttributes()
                      .HEADING;
                    if (para.getNextSibling()) {
                      var nextparastyle = para.getNextSibling()
                        .getAttributes();
                    }
                    if (p == 0 && paralength == 0 && nextparastyle != "Normal") {
                      para.removeFromParent();
                    }
                    if (parastyle == "Normal") {
                      if (prev) {
                        var prevparastyle = prev.getAttributes()
                          .HEADING;
                      }
                      if (prev && prev.getType() == "PARAGRAPH" && prevparastyle == "Normal") {
                        if (paralength > 0) {
                          para.getPreviousSibling()
                            .asParagraph()
                            .appendText(" ");
                        }
                        para.merge();
                      }
                    }
                  }
                }
              }
            }
          }
        }
      }
    }
    if (CLEAR_tabs == 'true') {
      removeTabs();
    }
    if (CLEAR_multiple == 'true') {
      removeMultipleSpaces();
    }
  }
  // No text selected
  else {
    DocumentApp.getUi()
      .alert('No text selected. Please select some text and try again.');
  }
}

function removeLinks() {
  var selection = DocumentApp.getActiveDocument()
    .getSelection();
  if (selection) {
    var elements = selection.getRangeElements();
    for (var i = 0; i < elements.length; i++) {
      var element = elements[i];
      // Only deal with text elements
      if (element.getElement()
        .editAsText) {
        var text = element.getElement()
          .editAsText();
        // Deal with partially selected text
        if (element.isPartial()) {
          text.setLinkUrl(element.getStartOffset(), element.getEndOffsetInclusive(), null);
        }
        // Deal with fully selected text
        else {
          text.setLinkUrl(null);
        }
      }
    }
  }
  // No text selected
  else {
    DocumentApp.getUi()
      .alert('No text selected. Please select some text and try again.');
  }
}

function removeLineBreaks() {
  var selection = DocumentApp.getActiveDocument()
    .getSelection();
  if (selection) {
    var elements = selection.getRangeElements();
    for (var i = 0; i < elements.length; i++) {
      var element = elements[i];
      // Only deal with text elements
      if (element.getElement()
        .editAsText) {
        var text = element.getElement()
          .editAsText();
        if (element.isPartial()) {
          var start = element.getStartOffset();
          var finish = element.getEndOffsetInclusive();
          var oldText = text.getText()
            .slice(start, finish);
          if (oldText.match(/\r/)) {
            var number = oldText.match(/\r/g)
              .length;
            for (var j = 0; j < number; j++) {
              var location = oldText.search(/\r/);
              text.deleteText(start + location, start + location);
              text.insertText(start + location, ' ');
              var oldText = oldText.replace(/\r/, ' ');
            }
          }
        }
        // Deal with fully selected text
        else {
          text.replaceText("\\v+", " ");
        }
      }
    }
  }
  // No text selected
  else {
    DocumentApp.getUi()
      .alert('No text selected. Please select some text and try again.');
  }
}

function removeParaBreaks() {
  var selection = DocumentApp.getActiveDocument()
    .getSelection();
  if (selection) {
    var elements = selection.getRangeElements();
    for (var i = 0; i < elements.length; i++) {
      var element = elements[i];
      // Only deal with text elements
      if (element.getElement()
        .editAsText) {
        var text = element.getElement()
          .editAsText();
        if (element.isPartial()) {
          var type = element.getElement()
            .getParent()
            .getType();
          if (type == "PARAGRAPH") {
            var para = element.getElement()
              .getParent()
            var prev = para.getPreviousSibling();
            var paratext = para.asText()
              .getText();
            var paralength = paratext.length;
            var finalchar = paratext.charAt(paralength - 1);
            if (paralength > 0 && finalchar == " ") {
              text.deleteText(paralength - 1, paralength - 1);
            }
            var parastyle = para.getAttributes()
              .HEADING;
            if (i == 0 && paralength == 0 && nextparastyle != "Normal") {
              para.removeFromParent();
            };
            if (parastyle == "Normal") {
              if (prev) {
                var prevparastyle = prev.getAttributes()
                  .HEADING
              }
              if (i > 0 && prev.getType() == "PARAGRAPH" && prevparastyle == "Normal") {
                para.getPreviousSibling()
                  .asParagraph()
                  .appendText(" ");
                para.merge();
              }
            }
          }
        }
        // Deal with fully selected text
        else {
          var type = element.getElement()
            .getType();
          if (type == "PARAGRAPH") {
            var para = element.getElement();
            var prev = para.getPreviousSibling();
            var paratext = para.asText()
              .getText();
            var paralength = paratext.length;
            var finalchar = paratext.charAt(paralength - 1);
            if (paralength > 0 && finalchar == " ") {
              text.deleteText(paralength - 1, paralength - 1);
            }
            var parastyle = para.getAttributes()
              .HEADING;
            if (para.getNextSibling()) {
              var nextparastyle = para.getNextSibling()
                .getAttributes();
            }
            if (i == 0 && paralength == 0 && nextparastyle != "Normal") {
              para.removeFromParent();
            }
            if (parastyle == "Normal") {
              if (prev) {
                var prevparastyle = prev.getAttributes()
                  .HEADING;
              }
              if (i > 0 && prev.getType() == "PARAGRAPH" && prevparastyle == "Normal") {
                if (paralength > 0) {
                  para.getPreviousSibling()
                    .asParagraph()
                    .appendText(" ");
                }
                para.merge();
              }
            }
          } else if (type == "TABLE_CELL") {
            var numchi = element.getElement()
              .asTableCell()
              .getNumChildren();
            for (var p = numchi - 1; p >= 0; p--) {
              var para = element.getElement()
                .asTableCell()
                .getChild(p);
              var celltext = para.editAsText();
              var prev = para.getPreviousSibling();
              var paratext = para.asText()
                .getText();
              var paralength = paratext.length;
              if (prev) {
                var prevtext = prev.asText()
                  .getText();
                var prevlength = prevtext.length;
                var prevchar = prevtext.charAt(prevlength - 1);
              }
              if (prev && prevlength > 0 && prevchar == " ") {
                prev.asText()
                  .deleteText(prevlength - 1, prevlength - 1);
              }
              var parastyle = para.getAttributes()
                .HEADING;
              if (para.getNextSibling()) {
                var nextparastyle = para.getNextSibling()
                  .getAttributes();
              }
              if (p == 0 && paralength == 0 && nextparastyle != "Normal") {
                para.removeFromParent();
              }
              if (parastyle == "Normal") {
                if (prev) {
                  var prevparastyle = prev.getAttributes()
                    .HEADING;
                  if (prev.getType() == "PARAGRAPH" && prevparastyle == "Normal") {
                    if (paralength > 0) {
                      para.getPreviousSibling()
                        .asParagraph()
                        .appendText(" ");
                    }
                  }
                  para.merge();
                }
              }
            }
          } else if (type == "TABLE") {
            var numrows = element.getElement()
              .asTable()
              .getNumChildren();
            for (var q = 0; q < numrows; q++) {
              var row = element.getElement()
                .asTable()
                .getChild(q)
                .asTableRow();
              var numcells = row.getNumChildren();
              for (var r = 0; r < numcells; r++) {
                var cell = row.getChild(r)
                  .asTableCell();
                var numpara = cell.getNumChildren();
                for (var p = numpara - 1; p >= 0; p--) {
                  var para = cell.asTableCell()
                    .getChild(p);
                  var celltext = para.editAsText();
                  var prev = para.getPreviousSibling();
                  var paratext = para.asText()
                    .getText();
                  var paralength = paratext.length;
                  if (prev) {
                    var prevtext = prev.asText()
                      .getText();
                    var prevlength = prevtext.length;
                    var prevchar = prevtext.charAt(prevlength - 1);
                  }
                  if (prev && prevlength > 0 && prevchar == " ") {
                    prev.asText()
                      .deleteText(prevlength - 1, prevlength - 1);
                  }
                  var parastyle = para.getAttributes()
                    .HEADING;
                  if (para.getNextSibling()) {
                    var nextparastyle = para.getNextSibling()
                      .getAttributes();
                  }
                  if (p == 0 && paralength == 0 && nextparastyle != "Normal") {
                    para.removeFromParent();
                  }
                  if (parastyle == "Normal") {
                    if (prev) {
                      var prevparastyle = prev.getAttributes()
                        .HEADING;
                      if (prev.getType() == "PARAGRAPH" && prevparastyle == "Normal") {
                        if (paralength > 0) {
                          para.getPreviousSibling()
                            .asParagraph()
                            .appendText(" ");
                        }
                      }
                      para.merge();
                    }
                  }
                }
              }
            }
          }
        }
      }
    }
  }
  // No text selected
  else {
    DocumentApp.getUi()
      .alert('No text selected. Please select some text and try again.');
  }
}

function removeBoth() {
  removeLineBreaks();
  removeParaBreaks();
}

function removeMultipleSpaces() {
  var selection = DocumentApp.getActiveDocument()
    .getSelection();
  if (selection) {
    var elements = selection.getRangeElements();
    for (var i = 0; i < elements.length; i++) {
      var element = elements[i];
      // Only deal with text elements
      if (element.getElement()
        .editAsText) {
        var text = element.getElement()
          .editAsText();
        if (element.isPartial()) {
          var start = element.getStartOffset();
          var finish = element.getEndOffsetInclusive();
          var oldText = text.getText()
            .slice(start, finish);
          if (oldText.match(/[ ][ ]+/g)) {
            var number = oldText.match(/[ ][ ]+/g)
              .length;
            for (var i = 0; i < number; i++) {
              var location = oldText.search(/[ ][ ]+/);
              var spaces = oldText.match(/[ ][ ]+/);
              text.deleteText(start + location, start + location + spaces[0].length - 2);
              var oldText = oldText.replace(/[ ][ ]+/, ' ');
            }
          }
        }
        // Deal with fully selected text
        else {
          text.replaceText("[ ][ ]+", " ");
        }
      }
    }
  }
  // No text selected
  else {
    DocumentApp.getUi()
      .alert('No text selected. Please select some text and try again.');
  }
}

function removeTabs() {
    var selection = DocumentApp.getActiveDocument()
      .getSelection();
    if (selection) {
      var elements = selection.getRangeElements();
      for (var i = 0; i < elements.length; i++) {
        var element = elements[i];
        // Only deal with text elements
        if (element.getElement()
          .editAsText) {
          var text = element.getElement()
            .editAsText();
          if (element.isPartial()) {
            var start = element.getStartOffset();
            var finish = element.getEndOffsetInclusive();
            var oldText = text.getText()
              .slice(start, finish);
            if (oldText.match(/\t/g)) {
              var number = oldText.match(/\t/g)
                .length;
              for (var i = 0; i < number; i++) {
                var location = oldText.search(/\t/);
                text.deleteText(start + location, start + location);
                text.insertText(start + location, ' ');
                var oldText = oldText.replace(/\t/, ' ');
              }
            }
          }
          // Deal with fully selected text
          else {
            text.replaceText("\\t", " ");
          }
        }
      }
    }
    // No text selected
    else {
      DocumentApp.getUi()
        .alert('No text selected. Please select some text and try again.');
    }
  }
  //----------------------------------------//
  // For testing purposes only

function clearProps() {
  var userProperties = PropertiesService.getUserProperties();
  userProperties.deleteAllProperties();
}
