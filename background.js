function isCompatible(version) {
  let [ major, minor , patch ] = version.split(".").map(e => parseInt(e,10));
  return (
    major > 102 || 
    (major == 102 && minor > 3) ||
    (major == 102 && minor == 3 && patch > 2)
  );
}

async function main() {
  let { version } = await browser.runtime.getBrowserInfo();
  if (isCompatible(version)) {
    await messenger.BootstrapLoader.registerChromeUrl([ ["content", "eas4tbsync", "content/"] ]);
    await messenger.BootstrapLoader.registerBootstrapScript("chrome://eas4tbsync/content/bootstrap.js");  
  } else {
    let manifest = browser.runtime.getManifest();
    browser.notifications.create({
      type: "basic",
      iconUrl: browser.runtime.getURL("content/skin/eas32.png"),
      title: `${manifest.name}`,
      message: "Please update Thunderbird to at least 102.3.3 to be able to use this provider.",
    });
  }
}

main();
