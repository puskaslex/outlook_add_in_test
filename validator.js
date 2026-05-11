const DEFAULT_ACTION = true; // true = Allow, false = block
const DEFAULT_MESSAGE = "This email was blocked due to a policy violation. Please review and modify the message before resending.";

var mailboxItem;
Office.initialize = function () {
    // This is mandatory, if this is not present it may cause the mailbox Object to be unpopulated at times
    // However this is not 100% guaranteed to work
    // improved version would be having an OnReady promise awaited form the event handler
    mailboxItem = Office.context.mailbox.item;
}

// Please note that this Reusable Web Socket class is intended to handle sequential
// communication req - response - req - response ...
// If ever there is a need to send multiple requests at once, this class needs changes
class ReusableWebSocket {
    constructor(url) {		
        this.url = url;
        this.socket = null;
        this.isOpen = false;
        this.messageQueue = [];		// Queue to store messages while WebSocket is closed
        this.pendingPromises = [];	// To handle messages that were sent directly
    }

    // Open the WebSocket connection
    open(timeout = 10000) { // timeout set to 10seconds (E.g: when client is not running)
        return new Promise((resolve, reject) => {
            const timeoutId = setTimeout(() => {
                if (!this.isOpen && this.socket) {
                        this.socket.close(); // Close the WebSocket if it's stuck
				}
                reject(new Error('WebSocket connection timed out.'));
            }, timeout);

            this.socket = new WebSocket(this.url);
            this.socket.binaryType = 'arraybuffer'; // Set WebSocket to handle binary data

            this.socket.addEventListener('open', () => {
                clearTimeout(timeoutId);
                this.isOpen = true;
               
				console.log('WebSocket connection opened');
				this._processQueue();
				
                resolve();
            });

            this.socket.addEventListener('message', (event) => {
                this._handleIncomingMessage(event.data); // Handle incoming messages
            });

            this.socket.addEventListener('close', (event) => {
                this.isOpen = false;
            });

            this.socket.addEventListener('error', (error) => {
                clearTimeout(timeoutId);
                reject(error);
                if (error.message && error.message.includes('certificate')) {
                    console.error('SSL Certificate Error. Check your certificate settings.');
                }
            });
        });
    }

    // Send a message and wait for a response
    send(message) {
        return new Promise((resolve, reject) => {
            if (this.isOpen) {
                // Send the message directly if the WebSocket is open
                this._sendMessageInternal(message, resolve, reject);
            } else {
                // Queue the message if the WebSocket is not open
                this.messageQueue.push({ message, resolve, reject });
            }
        });
    }

    // Internal method to send a message over the WebSocket
    _sendMessageInternal(message, resolve, reject) {
        try {
            this.socket.send(message); // Send the message (binary or text)
			this.pendingPromises.push({ resolve, reject });
        } catch (error) {
            reject(error); // If sending fails, reject the promise
        }
    }

	// Handle incoming WebSocket messages and resolve the corresponding promises
    _handleIncomingMessage(data) {
        if (this.pendingPromises.length > 0) {
            const { resolve, reject } = this.pendingPromises.shift(); // Pop the first item from the pending promises
            try {
                resolve(data); // Resolve the promise with the incoming data
            } catch (error) {
                reject(error); // Reject the promise if an error occurs while handling the data
            }
        } else {
            console.warn('Received message without a corresponding promise:', data);
        }
    }

    // Close the WebSocket connection
    close() {
        return new Promise((resolve, reject) => {
            if (this.isOpen) {
                
                this.socket.addEventListener('close', () => {
					this.isOpen = false;
					console.log('WebSocket closed successfully.');
					resolve();
                });
                
				this.socket.addEventListener('error', (error) => {
					this.isOpen = false;
					reject(new Error(`WebSocket error during close: ${error.message}`));
				});
                
                this.socket.close();
            } else {
                resolve();
            }
        });
    }
    
	// Process all queued messages once the WebSocket is open
    _processQueue() {
		if (!this.isOpen) {
			console.log('WebSocket is not open. Waiting for open...');
			return;
		}
    
        while (this.messageQueue.length > 0) {
            const { message, resolve, reject } = this.messageQueue.shift(); // Pop the first item from the queue
            this._sendMessageInternal(message, resolve, reject); // Send the message
        }
    }
}

async function onMessageSendHandler(event) {
    mailboxItem = Office.context.mailbox.item;
    const wsClient = new ReusableWebSocket('wss://localhost:27442/');

    try {
        await wsClient.open();

        let stop = await communicate(wsClient, event, "scan_check;");
        if (stop) return;

        const getSubject = async () => {
			try {
				return asyncGetValue(mailboxItem.subject.getAsync);
			} catch (error) {
				console.error('Failed to get subject:', error);
				return '[No Subject]'; // Provide a fallback subject
			}	
		};

        let recipients = await extractRecipients(mailboxItem);
        let subject = await getSubject();
        let message = "subject;" + recipients + subject;

        stop = await communicate(wsClient, event, message);
        if (stop) return;

		const getBody = async () =>  {
			try {
				return asyncGetValue(mailboxItem.body.getAsync.bind(mailboxItem.body, Office.CoercionType.Text));
			} catch (error) {
				console.error('Failed to get body:', error);
				return '[No Body]'; // Provide a fallback body
			}
		};
		
        let body = await getBody();
        console.log(body); 

        stop = await communicate(wsClient, event, "body;" + recipients + body);
        if (stop) return;

   
        const getAttachments = async () => asyncGetValue(mailboxItem.getAttachmentsAsync);
        const attachmentsInfo = await getAttachments();

        if (attachmentsInfo && attachmentsInfo.length > 0) {
            const attachmentPromises = attachmentsInfo.map(async (attachment) => {
                try {
                    // Await for the attachment content
                    const result = await getAttachmentContentAsync(attachment.id);

                    if (result.status === Office.AsyncResultStatus.Succeeded) {
                        const attachmentContent = result.value.content;
                        const attachmentFormat = result.value.format;  // "base64" or "arrayBuffer"
                        console.log(attachment.name)
                        console.log(attachmentFormat);
                        
						let stop = false;
                        if (attachmentFormat === "base64") {
							try {
								const arrayBuffer = base64ToArrayBuffer(attachmentContent);
								stop = await communicate(wsClient, event, combineStringAndArrayBuffer("attachment;" + recipients + attachment.name + ";", arrayBuffer)); 
							} catch (err) {
								console.error("Failed to process base64 attachment:", err);
								return;
							}
                        } else if (attachmentFormat === "arrayBuffer") {
							try {
								stop = await communicate(wsClient, event, combineStringAndArrayBuffer("attachment;" + recipients + attachment.name + ";", attachmentFormat));
							} catch (err) {
								console.error("Failed to process arrayBuffer attachment:", err);
								return;
							}
                        } else {
							console.error("Unsupported attachment format:", attachmentFormat);
						}
                        
                        if (stop) return;
                    } else {
                        console.error('Error retrieving attachment:', result.error);
                    }
                } catch (error) {
                    console.error('Error processing attachment:', error);
                }
            });
           
           await Promise.all(attachmentPromises);
        }
    } catch (error) { // Catches errors on both open and send
        // If the default action is to block then create an allowEvent variable and set it to false here!
        console.error('Error:', error);
    } finally {
        await wsClient.close(); // Gracefully close the WebSocket
        console.log('WebSocket connection closed.');
    }

    event.completed({ allowEvent: DEFAULT_ACTION });
}


// Helper function to wrap Office Async methods into promises
function asyncGetValue(getMethod) {
    return new Promise((resolve, reject) => {
        getMethod((result) => {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                resolve(result.value);
            } else {
                reject(result.error);
            }
        });
    });
}

async function extractRecipients(mailboxItem) {
	let recipients = "";
	
	try {
		if (mailboxItem.itemType === Office.MailboxEnums.ItemType.Appointment) {
			const [organizer, requiredAttendees, optionalAttendees] = await Promise.all([
					asyncGetValue(mailboxItem.organizer.getAsync),
					asyncGetValue(mailboxItem.requiredAttendees.getAsync),
					asyncGetValue(mailboxItem.optionalAttendees.getAsync)
				]);
				
			recipients = organizer.emailAddress;
			if (requiredAttendees.length > 0) {
				recipients += ";" + requiredAttendees.map(recipient => recipient.emailAddress).join(";");
			}
			if (optionalAttendees.length > 0) {
				recipients += ";" + optionalAttendees.map(recipient => recipient.emailAddress).join(";");
			}

			recipients += "|";  // Add separator for email recipients
		}
		else {
			 // For email messages, get from, to, cc, bcc recipients 
			const [from, to, cc, bcc ] = await Promise.all([
				asyncGetValue(mailboxItem.from.getAsync),
				asyncGetValue(mailboxItem.to.getAsync),
				asyncGetValue(mailboxItem.cc.getAsync),
				asyncGetValue(mailboxItem.bcc.getAsync)
			]);

			recipients = from.emailAddress + ";" + to.map(recipient => recipient.emailAddress).join(";");
			if (cc.length > 0) {
				recipients += ";" + cc.map(recipient => recipient.emailAddress).join(";");
			}
			if (bcc.length > 0) {
				recipients += ";" + bcc.map(recipient => recipient.emailAddress).join(";");
			}

			recipients += "|";  // Add separator for email recipients

		}
	} catch(error) {
		console.error("Error retrieving recipients:", error);
	}
    return recipients;
}

async function communicate(webSocket, evt, message)
{
    const response = await webSocket.send(message);
    const blockVar = "block";
    const allowVar = "allow";

    if (blockVar === response)
    {
		console.log('Blocking message...');
        await webSocket.close();
        mailboxItem.notificationMessages.addAsync('NoSend', {
            type: 'errorMessage',
            message: DEFAULT_MESSAGE
        });

        if (mailboxItem.isOnlineMeeting) //&& mailboxItem.itemType === Office.MailboxEnums.ItemType.Appointment)
        {
            await mailboxItem.removeAsync((result) => {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    console.log("Successfully removed the Teams meeting.");
                } else {
                    console.error("Failed to remove the Teams meeting.", result.error);
                }
            });
        }
        evt.completed({ allowEvent: false });
        return true;
    }
    else if (allowVar === response) 
    {
		console.log('Allowing message...');
        await webSocket.close();
        evt.completed({ allowEvent: true });
        return true;
    }
    return false;
}

// Function to wrap getAttachmentContentAsync in a promise
function getAttachmentContentAsync(attachmentId) {
    return new Promise((resolve, reject) => {
        mailboxItem.getAttachmentContentAsync(attachmentId, function (result) {
            if (result.status === Office.AsyncResultStatus.Succeeded) {
                resolve(result);
            } else {
				console.error(`Failed to get attachment content for ${attachmentId}:`, result.error);
                reject(result.error);
            }
        });
    });
}

// Function to convert Base64 to ArrayBuffer
function base64ToArrayBuffer(base64) {
    const binaryString = window.atob(base64);  // Decode base64 to binary string
    const len = binaryString.length;
    const bytes = new Uint8Array(len);
    for (let i = 0; i < len; i++) {
        bytes[i] = binaryString.charCodeAt(i);
    }
    return bytes.buffer;  // Return ArrayBuffer
}

function stringToUint8Array(str) {
    const encoder = new TextEncoder(); // TextEncoder encodes strings as UTF-8
    return encoder.encode(str);        // Returns Uint8Array
}

// Function to combine string (recipients info) and ArrayBuffer (attachment content)
function combineStringAndArrayBuffer(stringData, arrayBufferData) {
    const stringUint8Array = stringToUint8Array(stringData); // Convert string to Uint8Array
    const arrayBufferUint8Array = new Uint8Array(arrayBufferData); // Convert ArrayBuffer to Uint8Array if needed

    // Create a new Uint8Array large enough to hold both the string and the ArrayBuffer
    const combinedArray = new Uint8Array(stringUint8Array.length + arrayBufferUint8Array.length);

    // Copy the string data into the combined array
    combinedArray.set(stringUint8Array, 0); // Start at index 0
    // Copy the array buffer data into the combined array, after the string data
    combinedArray.set(arrayBufferUint8Array, stringUint8Array.length);

    // Return the combined data as an ArrayBuffer (for WebSocket transmission)
    return combinedArray.buffer;
}
