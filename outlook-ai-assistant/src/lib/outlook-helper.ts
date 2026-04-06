/**
 * Reads the body of the current email in read mode.
 * Returns the plain-text body of the email.
 */
export function getEmailBody(): Promise<string> {
  return new Promise((resolve, reject) => {
    const item = Office.context.mailbox.item;
    if (!item) {
      reject(new Error("No mail item is currently selected."));
      return;
    }
    item.body.getAsync(Office.CoercionType.Text, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
      } else {
        reject(new Error(result.error.message));
      }
    });
  });
}

/**
 * Reads the subject of the current email.
 */
export function getEmailSubject(): string {
  const item = Office.context.mailbox.item;
  if (!item) {
    throw new Error("No mail item is currently selected.");
  }
  return (item as Office.MessageRead).subject || "";
}

/**
 * Reads the sender display name and email address.
 */
export function getEmailSender(): { displayName: string; emailAddress: string } {
  const item = Office.context.mailbox.item as Office.MessageRead;
  if (!item || !item.from) {
    return { displayName: "", emailAddress: "" };
  }
  return {
    displayName: item.from.displayName || "",
    emailAddress: item.from.emailAddress || "",
  };
}

/**
 * Reads the body of the current draft in compose mode.
 */
export function getDraftBody(): Promise<string> {
  return new Promise((resolve, reject) => {
    const item = Office.context.mailbox.item;
    if (!item) {
      reject(new Error("No mail item is currently open."));
      return;
    }
    item.body.getAsync(Office.CoercionType.Text, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
      } else {
        reject(new Error(result.error.message));
      }
    });
  });
}

/**
 * Replaces the entire compose window body with the provided text.
 */
export function setDraftBody(text: string): Promise<void> {
  return new Promise((resolve, reject) => {
    const item = Office.context.mailbox.item as Office.MessageCompose;
    if (!item) {
      reject(new Error("No compose item is currently open."));
      return;
    }
    item.body.setAsync(text, { coercionType: Office.CoercionType.Text }, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve();
      } else {
        reject(new Error(result.error.message));
      }
    });
  });
}

/**
 * Inserts text at the current cursor position in a compose window.
 */
export function insertReply(text: string): Promise<void> {
  return new Promise((resolve, reject) => {
    const item = Office.context.mailbox.item as Office.MessageCompose;
    if (!item) {
      reject(new Error("No compose item is currently open."));
      return;
    }
    item.body.setSelectedDataAsync(
      text,
      { coercionType: Office.CoercionType.Text },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve();
        } else {
          reject(new Error(result.error.message));
        }
      }
    );
  });
}
