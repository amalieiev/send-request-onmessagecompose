import { debugMessage } from "./helpers/utils";
import jquery from "jquery";

Office.onReady(() => {
    // Call to initialise the Office components and enable the event based function
});

export async function onMessageComposeHandler(event: any): Promise<void> {
    try {
        const url = "https://amalieievsignatures.azurewebsites.net/api/test";
        let token: string = "Bearer qwerty123456";

        debugMessage("Start");

        try {
            debugMessage("Using fetch");
            await fetch(url, {
                method: "POST",
                mode: "cors",
                headers: {
                    "Content-Type": "application/json",
                    Authorization: `Bearer ${token}`,
                },
                body: JSON.stringify({
                    sender: "AdminJM@exclaimertest.com",
                }),
            });
            debugMessage("Using fetch Success");
        } catch (error: any) {
            debugMessage("Using fetch Error");
            debugMessage(error.message ? error.message : error);
        }

        try {
            debugMessage("Using jQuery ajax");
            await jquery.ajax({
                url: url,
                type: "POST",
                crossDomain: true,
                headers: {
                    "Content-Type": "application/json",
                    Authorization: `Bearer ${token}`,
                },
                data: JSON.stringify({
                    sender: "AdminJM@exclaimertest.com",
                }),
                dataType: "json",
            });
            debugMessage("Using jQuery ajax Success");
        } catch (error: any) {
            debugMessage("Using jQuery ajax Error");
            debugMessage(error.message ? error.message : error);
        }

        try {
            debugMessage("Using XMLHttpRequest");
            await new Promise((resolve, reject) => {
                const xhr = new XMLHttpRequest();
                xhr.open("POST", url, true);
                xhr.setRequestHeader("Content-Type", "application/json");
                xhr.setRequestHeader("Authorization", `Bearer ${token}`);
                xhr.withCredentials = true;
                xhr.timeout = 10000; // 10sec
                xhr.ontimeout = () => reject("Timeout Error");
                xhr.onload = () => {
                    if (xhr.status < 300) {
                        resolve(true);
                    } else {
                        reject(xhr.statusText);
                    }
                };
                xhr.onerror = () => reject(xhr.statusText);
                xhr.send(
                    JSON.stringify({
                        sender: "AdminJM@exclaimertest.com",
                    })
                );
            });
            debugMessage("Using XMLHttpRequest Success");
        } catch (error: any) {
            debugMessage("Using XMLHttpRequest Error");
            if (error) {
                debugMessage(error.message ? error.message : error);
            }
        }

        try {
            debugMessage("Using Azure Functions");
            await fetch(
                "https://amalieievsignatures.azurewebsites.net/api/test",
                {
                    method: "POST",
                    mode: "cors",
                    headers: {
                        "Content-Type": "text/plain",
                    },
                    body: JSON.stringify({
                        body: {
                            sender: "AdminJM@exclaimertest.com",
                        },
                        headers: {
                            Authorization: `Bearer ${token}`,
                        },
                    }),
                }
            );
            debugMessage("Using Azure Functions Success");
        } catch (error: any) {
            debugMessage("Using Azure Functions Error");
            debugMessage(error.message ? error.message : error);
        }
    } catch (error: any) {
        debugMessage(error.message ? error.message : error);
    }

    debugMessage("End");

    event.completed();
}

function getGlobal(): unknown {
    if (typeof self !== "undefined") {
        return self;
    }
    if (typeof window !== "undefined") {
        return window;
    }
    return typeof global !== "undefined" ? global : undefined;
}

const g = getGlobal() as any;

g.onMessageComposeHandler = onMessageComposeHandler;

Office.actions.associate("onMessageComposeHandler", onMessageComposeHandler);
