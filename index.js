function stockAlerts() {
    const IS_TEST_MODE_SETTING = "IsTestModeSetting";
    const EMAIL_RECIPIENT_SETTING = "EmailRecipientSetting";
    const PRICE_TARGET_RULES_RANGE = "PriceTargetRules";
    const ABOVE_TARGET_COMPARISON = "AboveTarget";
    const BELOW_TARGET_COMPARISON = "BelowTarget";

    const currentTime = Date.now();

    const isTestMode = SpreadsheetApp.getActiveSpreadsheet()
        .getRangeByName(IS_TEST_MODE_SETTING)
        .getValues()[0][0];

    const emailRecipient = SpreadsheetApp.getActiveSpreadsheet()
        .getRangeByName(EMAIL_RECIPIENT_SETTING)
        .getValues()[0][0];

    const priceTargetRules = SpreadsheetApp.getActiveSpreadsheet()
        .getRangeByName(PRICE_TARGET_RULES_RANGE)
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
            case ABOVE_TARGET_COMPARISON:
                if (currentPrice > targetPrice) {
                    sendEmail(
                        `Stock Alerts: ${ticker} ($${currentPrice}) above target`,
                        `${ticker} (${name}) is $${currentPrice} which is above $${targetPrice} (${currency}). Rule ID: ${ruleId}`
                    );
                }
                break;
            case BELOW_TARGET_COMPARISON:
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
