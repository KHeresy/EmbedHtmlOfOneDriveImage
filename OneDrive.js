// ==UserScript==
// @name         OneDrive Image Embed and Share
// @namespace    http://tampermonkey.net/
// @version      1.0
// @description  press 'e' to generate a embed HTML code ogf image in OneDrive
// @author       kheresy@gmail.com
// @match        https://onedrive.live.com/*
// @match        https://my.microsoftpersonalcontent.com/*
// @grant        none
// ==/UserScript==

(function() {
	'use strict';

	const uWaitTime = 100;
	let urlImage = '';

	function showMessage(msg) {
		let msgBox = document.getElementById('my-message-box');
		if (!msgBox) {
			msgBox = document.createElement('div');
			msgBox.id = 'my-message-box';
			msgBox.style.position = 'fixed';
			msgBox.style.top = '20px';
			msgBox.style.left = '50%';
			msgBox.style.transform = 'translateX(-50%)';
			msgBox.style.backgroundColor = '#333';
			msgBox.style.color = 'white';
			msgBox.style.padding = '10px 20px';
			msgBox.style.borderRadius = '10px';
			msgBox.style.zIndex = '9999';
			msgBox.style.fontSize = '16px';
			document.body.appendChild(msgBox);
		}
		msgBox.innerText = msg;
		msgBox.style.display = 'block';
		clearTimeout(msgBox._hideTimer);
		msgBox._hideTimer = setTimeout(() => {
			msgBox.style.display = 'none';
		}, 5000);
	}

	// step a: click ChevronDown
	function clickChevronDown() {
		const chevronIcon = document.querySelector('i[data-icon-name="ChevronDown"]');
		if (chevronIcon) {
			console.log('click permission selector');
			chevronIcon.click();

			// need to copy after click
			navigator.clipboard.readText().then(text => {
				urlImage = text.trim();
				if (urlImage) {
					console.log('embed URL: [' + urlImage + ']');
					navigator.clipboard.writeText('').then(() => {
						clickViewOnly()
					}).catch(err => {
						showMessage('ERROR: can not clear clipboard: ' + err);
					});
				} else {
					showMessage('ERROR: No embed URL in clipboard');
				}
			}).catch(err => {
				showMessage('ERROR: read from clipboard error: ' + err);
			});
		} else {
			console.log('Wait for share dialog');
			setTimeout(clickChevronDown, uWaitTime);
		}
	}

	// step b: change to view-only
	function clickViewOnly() {
		const viewOnly = document.querySelector("button.ms-ContextualMenu-link.root-85");
		if (viewOnly) {
			console.log('find view-only');
			viewOnly.click();

			setTimeout(() => {
				let copyButton = document.getElementById('copy-button');
				if (copyButton) {
					console.log('click copy button');
					copyButton.click();
					generateHtml();
				}
			}, uWaitTime);
		} else {
			setTimeout(clickViewOnly, uWaitTime);
		}
	}

	// step c: wait generated URL
	function generateHtml() {
		navigator.clipboard.readText().then(text => {
			const urlShare = text.trim();
			if (urlShare) {
				console.log('get share URL: ' + urlShare);
				const finalHtml = `<a target=_blank href="${urlShare}"><img src="${urlImage}"></a>`;
				navigator.clipboard.writeText(finalHtml).then(() => {
					showMessage('Copy share HTML: ' + finalHtml);
				}).catch(err => {
					showMessage('ERROR: Can NOT write into clipboard');
				});
			} else {
				setTimeout(generateHtml, uWaitTime);
			}
		}).catch(err => {
			showMessage('ERROR: read clipboard failed: ' + err);
		});
	}

	// step 1: find '...'
	function clickOverflowButton() {
		const button = document.querySelector('button[data-automationid="overflowButton"]');
		if (button) {
			console.log('click overflowButton');
			button.click();
			setTimeout(clickEmbedMenuItem, uWaitTime);
		} else {
			showMessage('can mpt find overflowButton');
		}
	}

	// step 2: click 'embed' in menu
	function clickEmbedMenuItem() {
		const menu = document.querySelector('ul.ms-ContextualMenu-list.is-open');
		if (menu) {
			const embedButton = menu.querySelector('button[data-automationid="embed"]');
			if (embedButton) {
				console.log('click embed in menu');
				embedButton.click();
				setTimeout(handleEmbedDetails, uWaitTime);
			} else {
				showMessage('can not find embed in menu');
			}
		} else {
			showMessage('can not fin menu');
		}
	}

	// step 3: process right panel
	function handleEmbedDetails() {
		const embedDetails = document.querySelector('.od-embed-details');
		if (!embedDetails) {
			console.log('can not find od-embed-detail, retrying');
			setTimeout(handleEmbedDetails, uWaitTime);
			return;
		}

		const generateButton = embedDetails.querySelector('button.od-embed-button');
		if (generateButton) {
			console.log('find and click generate button');
			generateButton.click();
			setTimeout(handleEmbedDetails, uWaitTime);
			return;
		}

		const sizePicker = embedDetails.querySelector('select.od-embed-sizePicker');
		
		if (sizePicker) {
			console.log('find image size selector');
			if (sizePicker.options.length > 0) {
				sizePicker.selectedIndex = sizePicker.options.length - 1;
				sizePicker.dispatchEvent(new Event('change', { bubbles: true }));
				console.log('set smallest size');
			}

			copyEmbedUrl();
		} else {
			showMessage('Can not find embed URL of image');
		}
	}

	// step 4: copy embed URL
	function copyEmbedUrl() {
		const textarea = document.querySelector('.od-embed-details').querySelector('textarea.od-TextField-field');
		const urlImage = textarea.value;
		if (urlImage) {
			console.log('find embed URL: ' + urlImage);

			navigator.clipboard.writeText(urlImage).then(() => {
				console.log('Copy the URL into clipboard for iframe script');
			}).catch(err => {
				showMessage('ERROR: can not write into clipboard: ' + err);
			});
			clickShareButton();
		} else {
			console.log('wait for URL generated');
			setTimeout(copyEmbedUrl, uWaitTime);
		}
	}

	// step 5: click share button
	function clickShareButton() {
		const shareButton = document.querySelector('button[data-automationid="share"]');
		if (shareButton) {
			console.log('click share, wait iframe process');
			shareButton.click();
		} else {
			showMessage('Can not find share button');
		}
	}

	// entrypoint
	if (location.hostname.includes('onedrive.live.com')) {
		window.addEventListener('keydown', function (event) {
			if (event.key === 'e' && !['INPUT', 'TEXTAREA'].includes(document.activeElement.tagName)) {
				clickOverflowButton();
			}
		});
	} else if (location.hostname.includes('microsoftpersonalcontent.com')) {
		console.log('In iframe page');
		clickChevronDown();
	}
})();
