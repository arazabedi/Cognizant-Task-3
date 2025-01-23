async function onFormLoad(executionContext) {
  const formContext = executionContext.getFormContext();
  const customerField = formContext.getAttribute("customerid");

  const primaryContactSection = formContext.ui.tabs
    .get("Summary")
    .sections.get("Summary_section_8");

  if (!customerField || !primaryContactSection) {
    console.warn("Field or section not found");
    return;
  }

	const customerInfo = getCustomerId(customerField);

  if (customerInfo) {
    await populateEmailAndPhone(formContext, customerInfo);
    primaryContactSection.setVisible(true);
  } else {
    primaryContactSection.setVisible(false);
  }
}

async function populateEmailAndPhone(formContext, customerInfo) {
  const { customerId, entityType } = customerInfo;

  // If it's an Account, get Primary Contact's details
  if (entityType === "account") {
    await getPrimaryContactDetails(formContext, customerId);
  } else {
    // If it's a Contact, get their own details
    await getContactDetails(formContext, customerId);
  }
}

function getCustomerId(customerField) {
  const customer = customerField.getValue()?.[0];
	if (customer) {
    // Clean GUID
    const customerId = customer.id.replace(/[{}]/g, "");
    // Get entity type (Account or Contact)
    const entityType = customer.entityType;
    return { customerId, entityType };
  }
  console.warn("No customer selected");
  return null;
}

async function getPrimaryContactDetails(formContext, accountId) {
  try {
    const account = await Xrm.WebApi.retrieveRecord(
      "account",
      accountId,
      "?$select=_primarycontactid_value"
    );

    const contactId = account["_primarycontactid_value"];
    if (contactId) {
      await getContactDetails(formContext, contactId);
    } else {
      console.warn("No primary contact for the account.");
    }
  } catch (err) {
    console.error("Error retrieving primary contact:", err);
  }
}

async function getContactDetails(formContext, contactId) {
  try {
    const contact = await Xrm.WebApi.retrieveRecord(
      "contact",
      contactId,
      "?$select=emailaddress1,mobilephone"
    );

    // Populate the email and phone number fields on the form
    setFieldValue(formContext, "emailaddress", contact.emailaddress1);
    setFieldValue(formContext, "cr4fd_mobilephonenumber", contact.mobilephone);
  } catch (err) {
    console.error("Error retrieving contact details:", err);
  }
}

function setFieldValue(formContext, fieldName, value) {
  const field = formContext.getAttribute(fieldName);
  if (field) {
    field.setValue(value);
  } else {
    console.warn(`Field '${fieldName}' not found on the form.`);
  }
}
