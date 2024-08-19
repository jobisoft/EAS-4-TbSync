async function main() {
  await messenger.BootstrapLoader.registerChromeUrl([ ["content", "eas4tbsync", "content/"] ]);
  await messenger.BootstrapLoader.registerBootstrapScript("chrome://eas4tbsync/content/bootstrap.js");  
}

main();
