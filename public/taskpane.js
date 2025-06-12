Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    checkRecipients();
  }
Recipients() {
  const item = Office.context.mailbox.item;
  const allowedDomain = "@nex-tone.co.jp";
  let warning = "";

  try {
    const toRecipients = item.to || [];
    const ccRecipients = item.cc || [];

    const allRecipients = [...toRecipients, ...ccRecipients];

    for (const recipient of allRecipients) {
      if (!recipient.emailAddress.endsWith(allowedDomain)) {
        warning = "⚠ 外部ドメインの宛先が含まれています: " + recipient.emailAddress;
        break;
      }
    }

    if (warning) {
      document.getElementById("warning").innerText = warning;
    }
  } catch (error) {
    console.error("Recipient check failed:", error);
  }
}
