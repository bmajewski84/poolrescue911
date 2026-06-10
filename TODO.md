# TODO

- **Lock down invoicing**: Add a shared passcode (Script Property, e.g. `PASSCODE`) so `invoice.html` and the
  `getInvoiceData` / `createInvoice` Apps Script actions require it. Currently the Web App is open to "Anyone",
  so anyone with the URL could generate Stripe payment links and trigger customer emails.
