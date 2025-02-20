/* eslint-disable no-undef */
/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

/* global global, Office, self, window */
Office.onReady(function (info) {
  if (info.host === Office.HostType.Excel) {
    // If needed, Office.js is ready to be called
  }
});

/**
 * Shows a notification when the add-in command is executed.
 * @param event
 */
function action(event: Office.AddinCommands.Event) {
  const message: Office.NotificationMessageDetails = {
    type: Office.MailboxEnums.ItemNotificationMessageType.InformationalMessage,
    message: "Performed action.",
    icon: "Icon.80x80",
    persistent: true,
  };
  // Show a notification message
  Office.context.mailbox.item.notificationMessages.replaceAsync("action", message);
  // Be sure to indicate when the add-in command function is complete
  event.completed();
}

async function highlightSelection(event) {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();
      range.format.fill.color = "yellow";
      await context.sync();
    });
  } catch (error) {
    console.error(error);
  }
  // Calling event.completed is required. event.completed lets the platform know that processing has completed.
  event.completed();
}
// You must register the function with the following line.
Office.actions.associate("highlightSelection", highlightSelection);

async function fertig(event) {
  try {
    await Excel.run(async (context) => {
      /**
       * Insert your Excel code here
       */
      const data = [["Fertig"]];
      const range = context.workbook.getSelectedRange();
      // Read the range address
      range.load("address");
      range.values = data;
      // Update the fill color
      range.format.fill.color = "#66ff00";
      // Update the Text color
      range.format.font.color = "black";
      // Update the Schrifttyp zu "Fett"
      range.format.font.bold = true;
      // Spaltenbreite automatisch anpassen
      range.format.autofitColumns();
      await context.sync();
      console.log(`The range address was ${range.address}.`);
    });
  } catch (error) {
    console.error(error);
  }
  // Calling event.completed is required. event.completed lets the platform know that processing has completed.
  event.completed();
}
// You must register the function with the following line.
Office.actions.associate("fertig", fertig);

async function addRow() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getItem("Täglicher Arbeitsplan");
    const expensesTable = sheet.tables.getItem("Arbeitsplan_Tabelle_Industriestr");

    expensesTable.rows.add(null, [["Linie", "Mitarbeiter", "Kunde", "Produkt"]]);

    sheet.getUsedRange().format.autofitColumns();
    sheet.getUsedRange().format.autofitRows();

    await context.sync();
  });
}
Office.actions.associate("addRow", addRow);

// The add-in command functions need to be available in global scope
function getGlobal() {
  return typeof self !== "undefined"
    ? self
    : typeof window !== "undefined"
    ? window
    : typeof global !== "undefined"
    ? global
    : undefined;
}

const g = getGlobal() as any;
g.action = action;
