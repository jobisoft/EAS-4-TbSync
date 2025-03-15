await browser.LegacyHelper.registerGlobalUrls([
  ["content", "eas4tbsync", "content/"],
]);

await browser.EAS4TbSync.load();
