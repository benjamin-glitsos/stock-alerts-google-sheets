function stockAlerts() {
    const currentTime = Date.now();

    const isTestMode = SpreadsheetApp.getActiveSpreadsheet()
        .getRangeByName("IsTestModeSetting")
        .getValues()[0][0];

    const emailRecipient = SpreadsheetApp.getActiveSpreadsheet()
        .getRangeByName("EmailRecipientSetting")
        .getValues()[0][0];

    const priceTargetRules = SpreadsheetApp.getActiveSpreadsheet()
        .getRangeByName("PriceTargetRules")
        .getValues();

    function sendEmail(subject, body) {
        if (isTestMode) {
            console.log("Send email ...", emailRecipient, subject, body);
        } else {
            MailApp.sendEmail(emailRecipient, subject, body, {
                noReply: true
            });
        }
    }

    for (const [
        ticker,
        isEnabled,
        comparison,
        targetPrice,
        startDate,
        endDate,
        ruleId,
        name,
        currentPrice,
        currency
    ] of priceTargetRules) {
        if (ticker === "") break;

        if (!isEnabled) continue;
        if (Date.parse(startDate) > currentTime) continue;
        if (Date.parse(endDate) < currentTime) continue;

        switch (comparison) {
            case "AboveTarget":
                if (currentPrice > targetPrice) {
                    sendEmail(
                        `Stock Alerts: ${ticker} ($${currentPrice}) above target`,
                        `${ticker} (${name}) is $${currentPrice} which is above $${targetPrice} (${currency}). Rule ID: ${ruleId}`
                    );
                }
                break;
            case "BelowTarget":
                if (currentPrice < targetPrice) {
                    sendEmail(
                        `Stock Alerts: ${ticker} ($${currentPrice}) below target`,
                        `${ticker} (${name}) is $${currentPrice} which is below $${targetPrice} (${currency}). Rule ID: ${ruleId}`
                    );
                }
                break;
        }
    }
}
