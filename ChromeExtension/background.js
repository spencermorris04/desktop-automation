import { startPolling } from './poll.js';

// run immediately on boot
startPolling();

// fire every minute in case the worker was frozen
chrome.alarms.create('keepAlive', { periodInMinutes: 1 });
chrome.alarms.onAlarm.addListener(a => a.name === 'keepAlive' && startPolling());
