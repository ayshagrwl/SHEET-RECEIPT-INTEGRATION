/*************************************************
 * RECEIPT ENTRY SYSTEM
 * Designed for existing Sales Voucher sheet
 *
 * Columns:
 * D = Invoice Number
 * E = Customer
 * F = Amount
 * G = Overdue
 * H = Discount
 * I = PAID-U
 * J = Status
 * K = MODE
 * L = OUTSTAND  <-- NEVER TOUCHED
 * M = RECEIPT
 * N = REMARKS
 *************************************************/


function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('Receipt Entry')
    .addItem('Add Receipt', 'openReceiptForm')
    .addToUi();
}


/**
 * Opens receipt popup for currently selected row
 */
function openReceiptForm() {

  const sheet = SpreadsheetApp.getActiveSheet();
  const row = sheet.getActiveCell().getRow();

  // Don't allow header row
  if (row <= 1) {
    SpreadsheetApp.getUi().alert(
      'Please select an invoice row first.'
    );
    return;
  }

  // Read existing values
  const invoice = sheet.getRange(row, 4).getDisplayValue(); // D
  const customer = sheet.getRange(row, 5).getDisplayValue(); // E
  const amount = sheet.getRange(row, 6).getDisplayValue(); // F
  const discount = sheet.getRange(row, 8).getDisplayValue(); // H
  const paid = sheet.getRange(row, 9).getDisplayValue(); // I
  const status = sheet.getRange(row, 10).getDisplayValue(); // J
  const mode = sheet.getRange(row, 11).getDisplayValue(); // K
  const outstanding = sheet.getRange(row, 12).getDisplayValue(); // L
  const receipt = sheet.getRange(row, 13).getDisplayValue(); // M
  const remarks = sheet.getRange(row, 14).getDisplayValue(); // N

  // Don't allow empty invoice rows
  if (!invoice) {
    SpreadsheetApp.getUi().alert(
      'The selected row does not contain an Invoice Number.'
    );
    return;
  }

  const template = HtmlService.createTemplateFromFile(
    'ReceiptForm'
  );

  template.row = row;
  template.invoice = invoice;
  template.customer = customer;
  template.amount = amount;
  template.discount = discount;
  template.paid = paid;
  template.status = status;
  template.mode = mode;
  template.outstanding = outstanding;
  template.receipt = receipt;
  template.remarks = remarks;

  const html = template.evaluate()
    .setWidth(470)
    .setHeight(720);

  SpreadsheetApp.getUi()
    .showModalDialog(html, 'Record Receipt');
}


/**
 * Saves receipt
 */
function saveReceipt(data) {

  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName(data.sheetName);

  if (!sheet) {
    throw new Error('Sales sheet not found.');
  }

  const row = Number(data.row);

  const cash = Number(data.cash) || 0;
  const bank = Number(data.bank) || 0;
  const discount = Number(data.discount) || 0;

  // At least one payment/discount required
  if (cash === 0 && bank === 0 && discount === 0) {
    throw new Error(
      'Please enter Cash, Bank, or Discount.'
    );
  }

  /********************************************
   * H = DISCOUNT
   ********************************************/
  if (discount !== 0) {

    const discountCell = sheet.getRange(row, 8);
    appendToExistingCell(
      discountCell,
      discount
    );
  }


  /********************************************
   * I = PAID-U
   *
   * IMPORTANT:
   * Existing formula is preserved.
   *
   * Example:
   *
   * =1000+600+800
   *
   * becomes:
   *
   * =1000+600+800+500+300
   ********************************************/
  const totalPayment = cash + bank;

  if (totalPayment !== 0) {

    const paidCell = sheet.getRange(row, 9);

    appendToExistingCell(
      paidCell,
      totalPayment
    );
  }


  /********************************************
   * J = STATUS
   *
   * We don't touch OUTSTAND.
   * We simply read its calculated result.
   ********************************************/

  SpreadsheetApp.flush();

  const outstandingCell = sheet.getRange(row, 12);

  const outstandingValue =
    Number(outstandingCell.getValue()) || 0;

  const statusCell = sheet.getRange(row, 10);

  if (outstandingValue <= 0) {
    statusCell.setValue('PAID');
  } else {
    statusCell.setValue('PARTIAL');
  }


  /********************************************
   * K = MODE
   *
   * Append:
   *
   * 11-Aug-2026 CASH ₹500 + BANK ₹300
   ********************************************/

  const modeCell = sheet.getRange(row, 11);

  const dateText = formatReceiptDate(
    data.date
  );

  const paymentParts = [];

  if (cash > 0) {
    paymentParts.push(
      'CASH ₹' + formatIndianNumber(cash)
    );
  }

  if (bank > 0) {
    paymentParts.push(
      'BANK ₹' + formatIndianNumber(bank)
    );
  }

  if (discount > 0) {
    paymentParts.push(
      'DISCOUNT ₹' + formatIndianNumber(discount)
    );
  }

  const modeEntry =
    dateText + ' ' + paymentParts.join(' + ');

  appendText(
    modeCell,
    modeEntry
  );


  /********************************************
   * M = RECEIPT
   *
   * Append receipt number.
   ********************************************/

  if (data.receiptNo) {

    const receiptCell =
      sheet.getRange(row, 13);

    appendText(
      receiptCell,
      data.receiptNo
    );
  }


  /********************************************
   * N = REMARKS
   *
   * Append new remarks.
   ********************************************/

  if (data.remarks) {

    const remarksCell =
      sheet.getRange(row, 14);

    appendText(
      remarksCell,
      data.remarks
    );
  }


  /********************************************
   * CREATE / UPDATE RECEIPTS SHEET
   ********************************************/

  let receiptSheet =
    ss.getSheetByName('Receipts');

  if (!receiptSheet) {

    receiptSheet =
      ss.insertSheet('Receipts');

    receiptSheet.appendRow([
      'Receipt Date',
      'Invoice Number',
      'Customer',
      'Invoice Amount',
      'Cash',
      'Bank',
      'Discount',
      'Total Payment',
      'Receipt Number',
      'Remarks',
      'Sales Sheet Row'
    ]);

    receiptSheet.setFrozenRows(1);
  }


  const invoice =
    sheet.getRange(row, 4).getDisplayValue();

  const customer =
    sheet.getRange(row, 5).getDisplayValue();

  const invoiceAmount =
    sheet.getRange(row, 6).getValue();


  receiptSheet.appendRow([
    data.date,
    invoice,
    customer,
    invoiceAmount,
    cash || '',
    bank || '',
    discount || '',
    totalPayment || '',
    data.receiptNo || '',
    data.remarks || '',
    row
  ]);


  return {
    success: true,
    outstanding: outstandingValue
  };
}


/**
 * Append numeric value while preserving
 * an existing formula.
 *
 * Existing:
 * =1000+600+800
 *
 * New amount:
 * 500
 *
 * Result:
 * =1000+600+800+500
 */
function appendToExistingCell(cell, value) {

  value = Number(value) || 0;

  if (value === 0) return;

  const formula = cell.getFormula();
  const currentValue = cell.getValue();

  // -----------------------------------------
  // CASE 1: Cell already contains a formula
  // -----------------------------------------
  if (formula) {

    cell.setFormula(
      formula + '+' + value
    );

    return;
  }


  // -----------------------------------------
  // CASE 2: Cell is completely blank
  // Create a NEW formula starting with =
  // -----------------------------------------
  if (
    currentValue === '' ||
    currentValue === null
  ) {

    cell.setFormula(
      '=' + value
    );

    return;
  }


  // -----------------------------------------
  // CASE 3: Cell contains a normal number
  // Convert it into a formula
  //
  // Example:
  // 1000
  //
  // becomes:
  // =1000+500
  // -----------------------------------------

  cell.setFormula(
    '=' + Number(currentValue) + '+' + value
  );
}


/**
 * Append text to existing cell
 *
 * Existing:
 * R3337
 *
 * New:
 * R4001
 *
 * Result:
 * R3337 + R4001
 */
function appendText(cell, newText) {

  if (!newText) return;

  const oldText =
    cell.getDisplayValue();

  if (!oldText) {

    cell.setValue(newText);

  } else {

    cell.setValue(
      oldText + ' + ' + newText
    );
  }
}


/**
 * Format date
 */
function formatReceiptDate(dateString) {

  if (!dateString) return '';

  const date =
    new Date(dateString);

  return Utilities.formatDate(
    date,
    Session.getScriptTimeZone(),
    'dd-MMM-yyyy'
  );
}


/**
 * Indian number formatting
 */
function formatIndianNumber(number) {

  return Number(number).toLocaleString(
    'en-IN',
    {
      maximumFractionDigits: 2
    }
  );
}
