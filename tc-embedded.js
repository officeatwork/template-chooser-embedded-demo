window.addEventListener("DOMContentLoaded", ignite);
window.addEventListener("message", handleEvent, false);

const defaultTemplateChooserDomain =
  "https://template-chooser-embedded.officeatwork.com";

let blobDocument = undefined;

let $template,
  $uploadUrl,
  $status,
  $spinner,
  $resultFile,
  $reload,
  $copyUrl,
  $tcInputUrl;

function handleEvent(event) {
  const templateChooserEmbeddedUrl =
    $tcInputUrl.value || defaultTemplateChooserDomain;
  if (!templateChooserEmbeddedUrl.includes(event.origin)) {
    return;
  }

  if (event.data && event.data.type === "template-chooser-status") {
    const { status } = event.data;
    $status.innerHTML = status;

    const showSpinner = status !== "Created";
    toggleSpinner(showSpinner);
  }

  if (event.data && event.data.type === "template-chooser-error") {
    const { error } = event.data;

    const status = `Error when creating document, error detail: <b>${JSON.stringify(
      error
    )}</b>`;
    $status.innerHTML = status;
    toggleSpinner(false);
  }

  if (event.data && event.data.type === "template-chooser-document-created") {
    const { blob, fileName } = event.data;

    blobDocument = blob;
    $resultFile.innerHTML = fileName;
    $resultFile.style.display = "inline-block";
    toggleSpinner(false);
  }

  if (event.data && event.data.type === "template-chooser-document-uploaded") {
    const { redirectUrl } = event.data;

    let status = "Document was uploaded";
    if (redirectUrl) {
      status += `. Response data: redirectUrl=<a href='${redirectUrl}' target='_blank'>${redirectUrl}</a>`;
    }

    $status.innerHTML = status;
    toggleSpinner(false);
  }
}

function ignite() {
  $template = document.querySelector("#template");
  $uploadUrl = document.querySelector("#upload-url");
  $status = document.querySelector("#status");
  $spinner = document.querySelector("#spinner");
  $resultFile = document.querySelector("#result-file");
  $reload = document.querySelector("#reload-tc");
  $copyUrl = document.querySelector("#copy-tc-url");
  $tcInputUrl = document.querySelector("#tc-input-url");
  $customXmlPart = document.querySelector("#customXmlPart");
  $insertSampleCustomXmlPart = document.querySelector(
    "#insert-sample-customxmlpart"
  );
  $tcInputUrl.value = defaultTemplateChooserDomain;

  $reload.addEventListener("click", () => {
    reloadIframe();
  });

  $copyUrl.addEventListener("click", () => {
    const embeddedUrl = document.getElementById(
      "template-chooser-embedded-iframe"
    ).src;
    window.navigator.clipboard.writeText(embeddedUrl);
  });

  $resultFile.addEventListener("click", () => {
    saveAs(blobDocument, $resultFile.innerHTML);
  });

  $insertSampleCustomXmlPart.addEventListener("click", () => {
    insertSampleCustomXmlPart();
  });

  function reloadIframe() {
    $status.innerHTML = "";
    $resultFile.innerHTML = "";
    $resultFile.style.display = "none";
    toggleSpinner(false);
    const $iframeContainer = document.getElementById("iframe-container");

    const iframe = document.createElement("iframe");
    iframe.src = buildEmbeddedUrl();
    iframe.id = "template-chooser-embedded-iframe";
    iframe.sandbox = "allow-same-origin allow-popups allow-scripts allow-forms";
    $iframeContainer.innerHTML = "";
    $iframeContainer.appendChild(iframe);
  }

  reloadIframe();
}

function toggleSpinner(show) {
  show ? ($spinner.style.display = "block") : ($spinner.style.display = "none");
}

function buildEmbeddedUrl() {
  const userInputUrl = $tcInputUrl.value || defaultTemplateChooserDomain;
  const [baseUrl, ...routing] = userInputUrl.split("#");

  let url = `${baseUrl}${buildEmbeddedParams()}`;

  if (routing && routing.length) {
    url += `#${routing.join("#")}`;
  }

  return url;
}

function buildEmbeddedParams() {
  const templateParam = buildTemplateParam();
  const injectParam = buildInjectParam();
  const uploadUrlParam = buildUploadUrlParam();

  let params = [templateParam, injectParam, uploadUrlParam]
    .filter((param) => !!param)
    .join("&");

  return params ? `?${params}` : "";
}

function buildTemplateParam() {
  const templateUrl = $template.value.trim();
  const base64Template = templateUrl.includes("template=")
    ? new URL(templateUrl).searchParams.get("template")
    : templateUrl || "";

  return base64Template ? `template=${base64Template}` : "";
}

function buildInjectParam() {
  const customXmlPart = $customXmlPart.value.trim();
  return customXmlPart ? `inject=${encodeCustomXmlPart(customXmlPart)}` : "";
}

function buildUploadUrlParam() {
  const uploadUrl = $uploadUrl.value.trim();
  return uploadUrl ? `uploadUrl=${uploadUrl}` : "";
}

function insertSampleCustomXmlPart() {
  const sampleCustomXmlPart = `<Properties xmlns="http://schemas.officeatwork.com/2022/templateProperties">
<officeatwork_languages>
  <Value>en</Value>
  <Value>de</Value>
</officeatwork_languages>
<subject>Invitation to branch opening</subject>
<subject.de>Einladung zur Geschäfstelleneröffnung</subject.de>
<location>
  <city>Zug</city>
  <country>Switzerland</country>
  <country.de>Schweiz</country.de>
  <street>Bundesplatz 12</street>
  <mapslink>https://www.bing.com/maps?osid=941010b2-4f60-4f55-9683-6ce7f9a2776b&amp;cp=47.172293~8.505079&amp;lvl=16&amp;imgid=aa3a3e86-b5f0-4c53-b3db-b348f9418b88&amp;v=2&amp;sV=2&amp;form=S00027</mapslink>
</location>
</Properties>`;

  $customXmlPart.value = sampleCustomXmlPart;
}

function encodeCustomXmlPart(customXmlPart) {
  const injectedData = [
    {
      type: "customxmlpart",
      base64Xml: encodeUnicodeToBase64(customXmlPart),
    },
  ];

  const encodedInjectedData = btoa(JSON.stringify(injectedData));

  return encodedInjectedData;
}

function encodeUnicodeToBase64(value) {
  return btoa(unescape(encodeURIComponent(value)));
}
