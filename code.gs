/*
 * 1. Menus and dialog
 * 2. User properties and settings
 * 3. Cleaning text: main function
 * 4. Paragraph iteration
 * 5. Individual cleaning functions
 */
 
function onInstall(e) {
    onOpen(e);
}

//
// # Menus and dialog
//

// Create Text Cleaner sub-menu
function onOpen(e) {
    DocumentApp.getUi()
        .createAddonMenu()
        .addItem('Clean selected text', 'cleanText')
        .addItem('Configure', 'showDialog')
        .addSeparator()
        .addItem('Remove links and underlining', 'menuLinks')
        .addSeparator()
        .addItem('Remove line breaks','menuLineBreaks')
        .addItem('Remove paragraph breaks', 'menuParagraphBreaks')
        .addItem('Replace double paragraph breaks', 'menuDoubleParagraphBreaks')
        .addSeparator()
        .addItem('Remove multiple spaces', 'menuMultipleSpaces')
        .addItem('Remove tabs', 'menuTabs')
        .addSeparator()
        .addItem('Smarten quotes', 'menuSmartenQuotes')
        .addToUi();
}

// Open configuration dialog
function showDialog() {
    var html = HtmlService
        .createHtmlOutputFromFile('dialog')
        .setSandboxMode(HtmlService.SandboxMode.IFRAME)
        .setWidth(360)
        .setHeight(260);
    DocumentApp.getUi().showModalDialog(html, 'Configure Text Cleaner');
}

//
// ## Quick menu functions
//

function menuLinks() {
    var to_execute = [removeLinks];
    iterateParagraphs(to_execute);
}


function menuLineBreaks() {
    var to_execute = [removeLineBreaks];
    iterateParagraphs(to_execute);
}


function menuParagraphBreaks() {
    var to_execute = [removeParagraphBreaks];
    iterateParagraphs(to_execute);
}


function menuMultipleSpaces() {
    var to_execute = [removeMultipleSpaces];
    iterateParagraphs(to_execute);
}


function menuTabs() {
    var to_execute = [removeTabs];
    iterateParagraphs(to_execute);
}


function menuSmartenQuotes() {
    var to_execute = [smartenQuotes];
    iterateParagraphs(to_execute);
}

function menuDoubleParagraphBreaks() {
    var to_execute = [replaceDoubleParagraphBreaks];
    iterateParagraphs(to_execute);
}

//
// # User properties and settings
//

function getOrSetProps(get_or_set, dialog_settings) {
    var user_properties = PropertiesService.getUserProperties(),
        the_properties = user_properties.getProperties(),
        setting_names = ['bold',
                         'italic',
                         'underline',
                         'strikethrough',
                         'indent',
                         'quotes',
                         'Links',
                         'LineBreaks',
                         'ParagraphBreaks',
                         'MultipleSpaces',
                         'Tabs'
                         ]
                         
    if (get_or_set === 'get') {
        var selected = (DocumentApp.getActiveDocument().getSelection()) ? true : false,
            user_settings = {'selected': selected};
        
        for (var i in setting_names) {
            var name = setting_names[i];
            user_settings[name] = the_properties['TXTCLN_' + name]
        }
        
        return user_settings;
    }
    if (get_or_set === 'set') {
        for (var i in setting_names) {
            var name = setting_names[i];
            user_properties.setProperty('TXTCLN_' + name, dialog_settings[i]);
        }
    }
}

//
// # Cleaning text: main function
//

function updateAndClean(dialog_settings) {
    // cleanText retrieves saved user properties, so set these first
    getOrSetProps('set', dialog_settings);
    cleanText();
}


function cleanText() {
    var user_properties = PropertiesService.getUserProperties().getProperties(),
        style = {},
        to_execute = [];
    
    // List of all the style attributes that are cleared as standard
    var std_clear = ['BACKGROUND_COLOR',
                     'BOLD',
                     'FONT_SIZE',
                     'FONT_FAMILY',
                     'FOREGROUND_COLOR',
                     'HORIZONTAL_ALIGNMENT',
                     'INDENT_FIRST_LINE',
                     'INDENT_END',
                     'INDENT_START',
                     'ITALIC',
                     'LINE_SPACING',
                     'SPACING_BEFORE',
                     'SPACING_AFTER',
                     'STRIKETHROUGH',
                     'UNDERLINE'
                     ]
                     
    // Check user properties to see whether a style attribute
    // is to be preserved, otherwise set it to null
    for (var i in std_clear) {
        var att_name = "TXTCLN_" + std_clear[i].split('_')[0].toLowerCase();
        if (user_properties[att_name] === 'checked') continue;
        style[eval("DocumentApp.Attribute." + std_clear[i])] = null;
    }
    
    // Check user properties to see whether a removal function should be run
    for (var i in user_properties) {
        var prop_name = i.substr(7, i.length),
            prop_val = user_properties[i];
        // Check only capitalized properties
        if (prop_name.substr(0, 1).toUpperCase() === prop_name.substr(0, 1)
            && prop_val === "checked") {
            // Eval must be used to convert string to function
            to_execute.push(eval('remove' + prop_name))
        }
    }
    
    // "Smarten quotes" is an exception
    var smarten = user_properties['TXTCLN_quotes'];
    if (smarten === "checked") to_execute.push(smartenQuotes);
    
    iterateParagraphs(to_execute, style);
}

//
// # Paragraph iteration
//

function iterateParagraphs(to_execute, style) {
    var selection = DocumentApp.getActiveDocument().getSelection();
    if (!selection) {
        DocumentApp.getUi().alert("No text selected." +
                                  "\nPlease select some text and try again");
        return
    }
    
    var elements = selection.getRangeElements();
    
    for (var i = 0; i < elements.length; i++) {
        var element = elements[i];
        // Skip the_elements that can't be edited as text (e.g images)
        if (!element.getElement().editAsText) continue;
        
        var the_element = element.getElement(),
            start_offset = (element.isPartial()) ? element.getStartOffset() : null,
            end_offset = (element.isPartial()) ? element.getEndOffsetInclusive() : null,
            element_type = the_element.getType();
            
        // Partially selected text
            
        if (start_offset != null) {
            if (style) clearFormatting(the_element, style, start_offset, end_offset);
            for (var process in to_execute) to_execute[process](the_element, i, start_offset, end_offset);
            continue
        }
        
        // Fully selected text (i.e. whole paragraphs)
        if (element_type != "TABLE" && element_type != "TABLE_CELL") {
            if (style) clearFormatting(the_element, style, start_offset);
            for (var process in to_execute) to_execute[process](the_element, i, start_offset);
        }
        
        // Tables
        
        // This function is never used again, so declared here
        function cleanTableCell(cell, style) {
            // Bury into the table cell to retrieve paragraphs
            for (var i = cell.getNumChildren() - 1; i >= 0; i--) {
                var the_element = cell.getChild(i);
                if (style) clearFormatting(the_element, style);
                for (var process in to_execute) to_execute[process](the_element, i);
            }
        }
        

        if (element_type == "TABLE_CELL") cleanTableCell(the_element.asTableCell(), style);
        
        // Bury into the table to retrieve cells
        if (element_type == "TABLE") {
            var table = the_element.asTable();
            for (var j = 0; j < table.getNumChildren(); j++) {
                var row = table.getChild(j).asTableRow();
                for (var k = 0; k < row.getNumChildren(); k++){
                    cleanTableCell(row.getChild(k), style);
                }
            }
        }
    }
}

//
// # Individual cleaning functions
//

// These operate at the text/paragraph level
// Doc Script's native .replace() method cannot be limited
// to parts of a text the_element. Partial text requires manual
// deletion of characters at an offset and manual insertion
// of replacement.

function removeLinks(the_element, paragraph_iteration, start_offset, end_offset) {
    if (start_offset == null) {
        the_element.setLinkUrl(null);
        return
    }
    the_element.setLinkUrl(start_offset, end_offset, null);
}


function removeTabs(the_element, paragraph_iteration, start_offset, end_offset) {
    if (start_offset == null) {
        the_element.replaceText("\\t+", " ");
        removeMultipleSpaces(the_element, paragraph_iteration, start_offset, end_offset);
        return
    }
  
    replaceCharInPartial(the_element, "\t", start_offset, end_offset);
  
    // Get rid of multiple spaces arising from multiple sequential
    // tabs being replaced by spaces.
    removeMultipleSpaces(the_element, paragraph_iteration, start_offset, end_offset);
}


function removeLineBreaks(the_element, paragraph_iteration, start_offset, end_offset) {
    if (start_offset == null) {
        the_element.replaceText("\\v+", " ");
        removeMultipleSpaces(the_element, paragraph_iteration, start_offset, end_offset);
        return
    }
  
    replaceCharInPartial(the_element, "\r", start_offset, end_offset);
    removeMultipleSpaces(the_element, paragraph_iteration, start_offset, end_offset);
}

// Manual replace characters in partial text -- used by removeLineBreaks and removeTabs
function replaceCharInPartial(the_element, char, start_offset, end_offset) {
    var current_text = the_element.getText().slice(start_offset, end_offset),
        reg = new RegExp(char),
        reg_global = new RegExp(char, "g");
        
    if (!current_text.match(reg_global)) return;
    
    var match_count = current_text.match(reg_global).length;
    for (var i = 0; i < match_count; i++) {
        var break_location = current_text.search(reg);
        the_element.deleteText(start_offset + break_location, start_offset + break_location)
                   .insertText(start_offset + break_location, " ");
      
        // Update the current_text object having deleted a character
        current_text = current_text.replace(reg, " ");
    }
}


function removeParagraphBreaks(the_element, paragraph_iteration, start_offset, end_offset) {
    var paragraph = (start_offset == null) ? the_element : the_element.getParent(),
        prev_paragraph = paragraph.getPreviousSibling();
        
    if (paragraph_iteration === 0) return;
    
    // Context for leaving paragraphs alone
    if (paragraph.getType() != "PARAGRAPH"
        || prev_paragraph.getType() != "PARAGRAPH"
        || paragraph.getAttributes().HEADING != "Normal"
        || prev_paragraph.getAttributes().HEADING != "Normal")
        {return};
    
    // Append space to preceding paragraph
    // only if this and that paragraph are not blank.
    if (paragraph.getText().length > 0
        && prev_paragraph.getText().length > 0)
        {prev_paragraph.appendText(" ")};
        
    paragraph.merge();
}


function removeMultipleSpaces(the_element, paragraph_iteration, start_offset, end_offset) {
    if (start_offset == null) {
        the_element.replaceText("[  ][  ]+", " ");
        return
    }
    
    var current_text = the_element.getText().slice(start_offset, end_offset);
    if (!current_text.match(/[  ][  ]+/g)) return;

    var match_count = current_text.match(/[  ][  ]+/g).length;
    for (var i = 0; i < match_count; i++) {
        var spaces_location = current_text.search(/[  ][  ]+/),
            the_spaces = current_text.match(/[  ][  ]+/);
            
        the_element.deleteText(start_offset + spaces_location,
                        start_offset + spaces_location
                        + the_spaces[0].length - 2);
                        
        current_text = current_text.replace(/[  ][  ]+/, " ");
    }
}


function smartenQuotes(the_element, paragraph_iteration, start_offset, end_offset) {
  
  // Smartening quotes in partial text is a nightmare, so isn't done
  // User is told in settings dialog that this only works with full paragraphs
  if (start_offset != null) return;
  
  var first_char = the_element.getText().charAt(0);
  
  if (first_char === '"') the_element.deleteText(0, 0).insertText(0, '“');
  if (first_char === "'") the_element.deleteText(0, 0).insertText(0, "’");
  
  the_element.replaceText(" '", " ‘")
         .replaceText(' "', " “")
         .replaceText("“'", "“‘")
         .replaceText('‘"', '‘“')
         .replaceText('"', '”')
         .replaceText("'", '’');
}


function clearFormatting(the_element, style, start_offset, end_offset) { 
    if (start_offset == null) {
    
        var indent =
            (the_element.getType() == 'LIST_ITEM')
            ? the_element.getIndentFirstLine()
            : style[DocumentApp.Attribute.INDENT_FIRST_LINE];
        
        the_element.setAttributes(style)
                   .setIndentFirstLine(indent);      
    } else {
        the_element.setAttributes(start_offset, end_offset, style);
    }
}


function replaceDoubleParagraphBreaks(the_element, paragraph_iteration) {
  
  
  // First, we deal with double paragraph breaks
  // If this works, don't look for double line breaks
  
  if (paragraph_iteration != 0 && the_element.getText().length > 0) {
      the_element.getPreviousSibling().appendText(' ');
      var new_paragraph = the_element.merge();
      Logger.log(new_paragraph);
      if (new_paragraph.getText().charAt(0) == ' ') new_paragraph.editAsText().deleteText(0,0);
      
  }
  
  
  // This code works for actual line breaks, but google doens't
  // paste plain text with these, so not really needed
  
  /*
  var new_paragraphs = the_element.getText().split(/\r{2,}/),
      doc = DocumentApp.getActiveDocument(),
      parent = the_element.getParent(),
      position = parent.getChildIndex(the_element),
      selection_range = doc.newRange(),
      new_text = (new_paragraphs[0].length > 0) ? new_paragraphs[0] : ' ';
  
  the_element.setText(new_text);
  selection_range.addElement(the_element);
  
  for (i = 1; i < new_paragraphs.length; i++) {
    var new_paragraph = parent.insertParagraph(position + i, new_paragraphs[i])
                              .setSpacingAfter(null);
    selection_range.addElement(new_paragraph);
  }
  
  doc.setSelection(selection_range);
  menuLineBreaks();
  */
}


//****Testing only****
function logProps() {
    Logger.log(PropertiesService.getUserProperties().getProperties());
}

function clearProps() {
    Logger.log(PropertiesService.getUserProperties().deleteAllProperties());
}
