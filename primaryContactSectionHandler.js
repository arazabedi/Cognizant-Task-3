// ===== Event Handler =====

// OnLoad event handler for the Primary Contact section
async function handlePrimaryContactSectionOnLoad(executionContext) {
  try {
    const formContext = executionContext.getFormContext();

    // Retrieve necessary fields and sections
    const customerField = formContext.getAttribute("customerid");
    const primaryContactSection = getPrimaryContactSection(formContext);

    if (!customerField || !primaryContactSection) {
      console.warn("Customer field or Primary Contact section not found");
      return;
    }

    // Process customer information
    const customerInfo = getCustomerId(customerField);
    if (customerInfo && customerInfo.entityType === "account") {
      await handleAccountCustomer(formContext, customerInfo);
      primaryContactSection.setVisible(true);
    }
  } catch (error) {
    console.error("Error handling Primary Contact section OnLoad:", error);
  }
}

// ===== Helper Functions =====

// Retrieve the Primary Contact section from the form
function getPrimaryContactSection(formContext) {
  const section = formContext.ui.tabs
    .get("Summary")
    .sections.get("Summary_section_8");
  return section || null;
}

// Process an account-type customer by populating email and phone fields
async function handleAccountCustomer(formContext, customerInfo) {
  try {
    const contactId = await getPrimaryContactId(customerInfo.customerId);
    if (contactId) {
      const contactDetails = await getContactDetails(contactId);
      updateFields(formContext, contactDetails);
    }
  } catch (error) {
    console.error("Error handling account customer:", error);
  }
}

// Update form fields with contact details
function updateFields(formContext, contactDetails) {
  if (contactDetails) {
    setFieldValue(formContext, "emailaddress", contactDetails.emailaddress1);
    setFieldValue(
      formContext,
      "cr4fd_mobilephonenumber",
      contactDetails.mobilephone
    );
  }
}

// Get the customer ID and entity type from the customer lookup
function getCustomerId(customerField) {
  const customer = customerField.getValue()?.[0];
  if (!customer) {
    console.warn("Customer not found");
    return null;
  }
  return {
    // Remove curly braces from the GUID
    customerId: customer.id.replace(/[{}]/g, ""),
    entityType: customer.entityType,
  };
}

// Set the value of a form field
function setFieldValue(formContext, fieldName, value) {
  const field = formContext.getAttribute(fieldName);
  if (field) {
    field.setValue(value);
  } else {
    console.warn(`Field '${fieldName}' not found on the form.`);
  }
}

// ===== API Calls =====

// Retrieve the primary contact ID for an account
async function getPrimaryContactId(accountId) {
  try {
    const account = await Xrm.WebApi.retrieveRecord(
      "account",
      accountId,
      "?$select=_primarycontactid_value"
    );
    return account["_primarycontactid_value"] || null;
  } catch (error) {
    console.error("Error retrieving primary contact ID:", error);
    return null;
  }
}

// Retrieve the details of a contact
async function getContactDetails(contactId) {
  try {
    const contact = await Xrm.WebApi.retrieveRecord("contact", contactId);
    return contact;
  } catch (error) {
    console.error("Error retrieving contact details:", error);
    return null;
  }
}
