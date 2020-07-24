function handleUpdateAvailable(details) {
  console.log("Update available for Eas4TbSync");
}

async function main() {
  // just by registering this listener, updates will not install until next restart
  //messenger.runtime.onUpdateAvailable.addListener(handleUpdateAvailable);

  await messenger.BootstrapLoader.registerChromeUrl([ ["content", "eas4tbsync", "content/"] ]);
  await messenger.BootstrapLoader.registerBootstrapScript("chrome://eas4tbsync/content/bootstrap.js");  
}

main();
