# Text Cleaner
Text Cleaner is a free add-on for Google Docs that adds powerful and customisable formatting clearance.

# Installation

Text Cleaner is available free of charge in the Google Add-ons store, which can be accessed from within a Google document by going to the 'Add-ons' menu and selecting 'Get add-ons...' or by clicking the link above. Once installed, Text Cleaner will appear in the 'Add-ons' menu.

# Clean selected text

This action applies the user's chosen settings to the selected text. Cleaning text can sometimes take a little while because the scripting that underlies the 'remove' functions is complicated. It tests for a number of conditions within each selected element and reacts accordingly. For very large documents, this can be *a lot* of work for the script (see below for advice concerning configuration for large documents). Disabling these functions when they are not needed speeds things up.

# Configure

The configuration dialog provides check boxes under two categories: 'preserve' and 'remove'. All text and paragraph attributes other than those listed under 'preserve' are cleared from selected text when text is cleaned. Listed below is the formatting that is removed by default by Text Cleaner (i.e. cannot be preserved). **Starred items** apply only to paragraphs and **will not be cleared unless an entire paragraph is selected** (e.g. by triple-clicking anywhere in a paragraph). This means that it is actually possible to preserve starred items provided the paragraph is not entirely selected (e.g. by adding a space at the end of the paragraph and then selecting all but this space). Cleared attributes:

+ Font
+ Font size
+ Text colour
+ Text background colour
+ Horizontal alignment (centre, left, right, justify)*
+ Line spacing*
+ Spacing before and after paragraph*
+ Non-standard list indentation*
+ Non-standard spacing after bullet points or numbers in lists*

**Important:** Text Cleaner is designed to be able to deal with text copied from outside Google Docs. Often the line breaks present in such text, even those that were not paragraph breaks in the original text, are interpreted as such by Docs. As a result, the 'paragraph breaks' option sometimes needs to be selected in order to remove line breaks from copied text. **There is no way to tell which paragraph breaks are converted line breaks**, which means that selecting this option **will remove** genuine paragraph breaks. If you want to preserve paragraphs while removing these converted line breaks, you will need to clear paragraphs individually. This cannot be corrected, since the information required to script for this is simply not present for the script to work with.

Text Cleaner stores settings for the user whenever they are changed. This means that the settings are always those that were last used, even when opening a different document or creating a new document.

As I said above, the 'remove' functions are based on complicated scripting and can be disabled when not required in order to speed up the add-on. If you need to clear a very long document, it is best to run functions separately. That is, deselect all of the 'remove' options and clean text. Then run the separate remove functions from the Text Cleaner menu. If you need to remove line breaks, but not paragraph breaks, then do not deselect the option to remove line breaks on the first clean, since there are not separate remove functions for line breaks and paragraph breaks in the Text Cleaner menu.

Removing links necessarily removes underlining in selected text, even if the underlining is not the result of a link. For this reason, Text Cleaner does not allow you to choose both to preserve underlining and to remove links. Checking the box to remove links deselects and disables the box to preserve underlining. Google's app scripting language treats the command to remove links as tantamount to 'remove links and all style attributes associated with them, including underlining'. This removal of underlining will be applied to all selected text. Sorry!

# Removal buttons

The Text Cleaner menu provides buttons to quickly remove links and underlining; line and paragraph breaks; multiple spaces; and tabs. These options have been included because sometimes the user wants to preserve most of the formatting, but bring paragraph text together properly.


