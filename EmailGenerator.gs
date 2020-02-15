// JSHint - TODO
/* jshint asi: true */

(function() {"use strict"})()

// EmailGeneratorgs
// =================
//
// External interface to this script - all of the event handlers.
//
// This files contains all of the event handlers, plus miscellaneous functions 
// not worthy of their own files yet
//
// The filename is prepended with _API as the Github chrome extension won't 
// push a file with the same name as the project.

var Log_

// Public event handlers
// ---------------------
//
// All external event handlers need to be top-level function calls; they can't 
// be part of an object, and to ensure they are all processed similarily 
// for things like logging and error handling, they all go through 
// errorHandler_(). These can be called from custom menus, web apps, 
// triggers, etc
// 
// The main functionality of a call is in a function with the same name but 
// post-fixed with an underscore (to indicate it is private to the script)
//
// For debug, rather than production builds, lower level functions are exposed
// in the menu

var EVENT_HANDLERS_ = {

//                           Name                            onError Message                          Main Functionality
//                           ----                            ---------------                          ------------------

  generateHtmlEmail:         ['generateHtmlEmail()',         'Failed to generate email',              generateHtmlEmail_],
  addNewFieldsForInput:      ['addNewFieldsForInput()',      'Failed to addNewFieldsForInput',        addNewFieldsForInput_],
  hideEmptyRows:             ['hideEmptyRows()',             'Failed to hideEmptyRows',               hideEmptyRows_],
  showAllRows:               ['showAllRows()',               'Failed to showAllRows',                 showAllRows_],
  reformatSpreadsheet:       ['reformatSpreadsheet()',       'Failed to reformatSpreadsheet',         reformatSpreadsheet_],
  hideOldColumns:            ['hideOldColumns()',            'Failed to hideOldColumns',              hideOldColumns_],
  showAllColumns:            ['showAllColumns()',            'Failed to showAllColumns',              showAllColumns_],
  removeEmptyColumns:        ['removeEmptyColumns()',        'Failed to removeEmptyColumns',          removeEmptyColumns_],
  archiveCurrentColumn:      ['archiveCurrentColumn()',      'Failed to archiveCurrentColumn',        archiveCurrentColumn_],
}

function generateHtmlEmail(args)    {return eventHandler_(EVENT_HANDLERS_.generateHtmlEmail, args)}
function addNewFieldsForInput(args) {return eventHandler_(EVENT_HANDLERS_.addNewFieldsForInput, args)}
function hideEmptyRows(args)        {return eventHandler_(EVENT_HANDLERS_.hideEmptyRows, args)}
function showAllRows(args)          {return eventHandler_(EVENT_HANDLERS_.showAllRows, args)}
function reformatSpreadsheet(args)  {return eventHandler_(EVENT_HANDLERS_.reformatSpreadsheet, args)}
function hideOldColumns(args)       {return eventHandler_(EVENT_HANDLERS_.hideOldColumns, args)}
function showAllColumns(args)       {return eventHandler_(EVENT_HANDLERS_.showAllColumns, args)}
function removeEmptyColumns(args)   {return eventHandler_(EVENT_HANDLERS_.removeEmptyColumns, args)}
function archiveCurrentColumn(args) {return eventHandler_(EVENT_HANDLERS_.archiveCurrentColumn, args)}

// Private Functions
// =================

// General
// -------

/**
 * All external function calls should call this to ensure standard 
 * processing - logging, errors, etc - is always done.
 *
 * @param {Array} config:
 *   [0] {Function} prefunction
 *   [1] {String} eventName
 *   [2] {String} onErrorMessage
 *   [3] {Function} mainFunction
 *
 * @param {Object}   args       The argument passed to the top-level event handler
 */

function eventHandler_(config, args) {

  try {

    var userEmail = Session.getActiveUser().getEmail()

    Log_ = BBLog.getLog({
      level:                DEBUG_LOG_LEVEL_, 
      displayFunctionNames: DEBUG_LOG_DISPLAY_FUNCTION_NAMES_,
    })
    
    Log_.info('Handling ' + config[0] + ' from ' + (userEmail || 'unknown email') + ' (' + SCRIPT_NAME + ' ' + SCRIPT_VERSION + ')')
    
    // Call the main function
    return config[2](args)
    
  } catch (error) {
  
    var assertConfig = {
      error:          error,
      userMessage:    config[1],
      log:            Log_,
      handleError:    HANDLE_ERROR_, 
      sendErrorEmail: SEND_ERROR_EMAIL_, 
      emailAddress:   ADMIN_EMAIL_ADDRESS_,
      scriptName:     SCRIPT_NAME,
      scriptVersion:  SCRIPT_VERSION, 
    }

    Assert.handleError(assertConfig) 
  }
  
} // eventHandler_()
