/**
 * Provides the platform on which the add-in is running.
 *
 * @param platform - Office.PlatformType
 *
 * **Important**: In Outlook, Office.context.platform property is available from requirement set 1.5.
 * For all Mailbox requirement sets, you can use the `Office.context.diagnostics` property to get the platform.
 */
export const isPlatform = (platform: Office.PlatformType): boolean => {
    try {
        return Office.context.platform === platform;
    } catch (error) {
        return Office.context.diagnostics.platform === platform;
    }
};

const messages: string[] = [];

export async function debugMessage(text: string): Promise<void> {
    messages.push(text);

    const messageHTML = `
		<table>
			${messages
                .map((message) => {
                    return `<tr><td>${message}</td></tr>`;
                })
                .join("")}
		</table>
	`;

    return new Promise((resolve) => {
        Office.context.mailbox.item?.body.setSignatureAsync(
            messageHTML,
            { coercionType: Office.CoercionType.Html },
            () => {
                resolve();
            }
        );
    });
}
