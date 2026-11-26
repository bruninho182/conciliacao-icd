chrome.action.onClicked.addListener(async (tab) => {
    if (!tab.id || new URL(tab.url || 'about:blank').href.split(/[?#]/)[0].replace(/\/$/, '') !==
        'https://app.icdgrupo.com.br/apps/sales/bookings') {
        await chrome.tabs.create({url: 'https://app.icdgrupo.com.br/apps/sales/bookings'});
        return;
    }
    await chrome.scripting.executeScript({target: {tabId: tab.id}, files: ['xlsx.js', 'painel.js']});
});
