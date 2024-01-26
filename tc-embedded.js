window.addEventListener("DOMContentLoaded", ignite);
window.addEventListener("message", handleEvent, false);

const defaultTemplateChooserDomain =
  "https://template-chooser-embedded.officeatwork.com";

let blobDocument = undefined;

let $template,
  $uploadUrl,
  $uploadClientId,
  $spinner,
  $result,
  $reload,
  $copyUrl,
  $tcInputUrl;

function handleEvent(event) {
  const templateChooserEmbeddedUrl =
    $tcInputUrl.value || defaultTemplateChooserDomain;
  if (!templateChooserEmbeddedUrl.includes(event.origin)) {
    return;
  }

  if (event.data) {
    event.data.type === "template-chooser-error" && setErrorStatus(event);

    event.data.type === "template-chooser-document-created" &&
      setCreatedResultFile(event);

    event.data.type === "template-chooser-document-uploaded" &&
      setUploadedStatus(event);

    event.data.type === "template-chooser-template-chosen" &&
      setTemplateChosenResult(event);
  }
}

function setErrorStatus(event) {
  const result = `Error when creating document, error detail: <pre><code>${JSON.stringify(
    event.data.error,
    null,
    4
  )}</code></pre>`;

  setResult(result);
}

function setCreatedResultFile(event) {
  const { blob, fileName } = event.data;

  const result = `<a id='created-file' href="JavaScript:void(0);">${fileName}</a>`;
  setResult(result, blob);
}

function setUploadedStatus(event) {
  const { redirectUrl } = event.data;

  let result = `Document was uploaded at ${new Date().toLocaleTimeString()} <br/>`;
  if (redirectUrl) {
    result += `Response data: redirectUrl=<a href='${redirectUrl}' target='_blank'>${redirectUrl}</a>`;
  }

  setResult(result);
}

function setTemplateChosenResult(event) {
  const { template: deepLink } = event.data;
  const result = `Link created at ${new Date().toLocaleTimeString()} <a href='${deepLink}' target='_blank'>Click here</a>`;
  setResult(result);
}

function setResult(result, blob) {
  blobDocument = blob;
  $result.innerHTML = result;
}

function ignite() {
  $template = document.querySelector("#template");
  $uploadUrl = document.querySelector("#upload-url");
  $uploadClientId = document.querySelector("#upload-client-id");
  $result = document.querySelector("#result");
  $reload = document.querySelector("#reload-tc");
  $copyUrl = document.querySelector("#copy-tc-url");
  $tcInputUrl = document.querySelector("#tc-input-url");
  $customXmlPart = document.querySelector("#customXmlPart");
  $documentProperties = document.querySelector("#documentProperties");
  $insertSampleCustomXmlPart = document.querySelector(
    "#insert-sample-customxmlpart"
  );
  $insertSampleDocumentProperties = document.querySelector(
    "#insert-sample-documentproperties"
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

  $result.addEventListener("click", () => {
    const templateLink$ = $result.querySelector("#created-file");
    if (templateLink$) {
      saveAs(blobDocument, templateLink$.innerHTML);
    }
  });

  $insertSampleCustomXmlPart.addEventListener("click", () => {
    insertSampleCustomXmlPart();
  });
  $insertSampleDocumentProperties.addEventListener("click", () => {
    insertSampleDocumentProperties();
  });

  function reloadIframe() {
    $result.innerHTML = "";
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

function buildEmbeddedUrl() {
  const userInputUrl = $tcInputUrl.value || defaultTemplateChooserDomain;
  const [baseUrlWithParams, hash] = userInputUrl.split("#");
  const [baseUrl, initialParams] = baseUrlWithParams.split("?");

  const hashParams = buildHashParams(hash);
  const queryParams = buildQueryParams(initialParams, hashParams);

  const url = `${baseUrl}${queryParams}${hashParams}`;
  return url;
}

function buildQueryParams(initialParams, hashParams) {
  if (
    hashParams.includes("uploadUrl") ||
    hashParams.includes("uploadClientId")
  ) {
    return "";
  }
  const uploadUrlParam = buildUploadUrlParam();
  const uploadClientId = buildUploadClientIdParam();

  let params = [initialParams, uploadUrlParam, uploadClientId]
    .filter((param) => !!param)
    .join("&");

  return params ? `?${params}` : "";
}

function buildHashParams(hash) {
  const initialHash = hash ? hash.trim() : "";
  const templateParam = buildTemplateParam();
  const injectParam = buildInjectParam();

  if (!initialHash && !templateParam && !injectParam) {
    return "";
  }

  const uploadUrlParam = buildUploadUrlParam();
  const uploadClientIdParam = buildUploadClientIdParam();

  const params = [
    initialHash,
    templateParam,
    injectParam,
    uploadUrlParam,
    uploadClientIdParam,
  ]
    .filter((param) => !!param)
    .join("&");

  return params ? `#${params}` : "";
}

function buildTemplateParam() {
  const urlOrBase64 = $template.value.trim();
  if (!urlOrBase64) {
    return "";
  }

  const templateParam = "template";
  const isBase64Template = !urlOrBase64.includes(`${templateParam}=`);

  if (isBase64Template) {
    return `${templateParam}=${urlOrBase64}`;
  }

  const urlTemplate = new URL(urlOrBase64);
  const queryOrHash = (urlTemplate.search || urlTemplate.hash).substring(1);
  const base64Template = new URLSearchParams(queryOrHash).get(templateParam);

  return base64Template ? `${templateParam}=${base64Template}` : "";
}

function buildUploadUrlParam() {
  const uploadUrl = $uploadUrl.value.trim();
  return uploadUrl ? `uploadUrl=${encodeURIComponent(uploadUrl)}` : "";
}

function buildUploadClientIdParam() {
  const uploadClientId = $uploadClientId.value.trim();
  return uploadClientId ? `uploadClientId=${uploadClientId}` : "";
}

function buildInjectParam() {
  const customXmlPart = $customXmlPart.value.trim();
  const documentProperties = $documentProperties.value.trim();
  if (!customXmlPart && !documentProperties) {
    return "";
  }
  return `inject=${prepareInjectionData(customXmlPart, documentProperties)}`;
}

function insertSampleCustomXmlPart() {
  const sampleCustomXmlPart = `<Properties xmlns="http://schemas.officeatwork.com/2022/templateProperties">
<officeatwork_languages>
  <Value>en</Value>
  <Value>de</Value>
</officeatwork_languages>
<subject>Invitation to branch opening</subject>
<subject.de>Einladung zur Geschäftsstelleneröffnung</subject.de>
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

function insertSampleDocumentProperties() {
  const sampleDocumentProperties = `{
    "type": "customdocumentproperty",
    "name": "MyProperty",
    "value": "Some Value"
 },
 {
    "type": "customdocumentproperty",
    "name": "My2ndProperty",
    "value": "Some Other Value"
 },
 {
    "type": "builtindocumentproperty",
    "name": "Title",
    "value": "Some Title"
 }`;

  $documentProperties.value = sampleDocumentProperties;
}

function prepareInjectionData(customXmlPart, documentProperties) {
  let injectedData = [];

  if (customXmlPart) {
    injectedData.push({
      type: "customxmlpart",
      base64Xml: encodeUnicodeToBase64(customXmlPart),
    })
  }

  if (documentProperties) {
    try {
      const documentPropertiesData = JSON.parse(
        `[${documentProperties}]`
      );

      injectedData = injectedData.concat(documentPropertiesData);
    } catch {
      // ignore invalid string
    }
  }

  const base64EncodedData = btoa(JSON.stringify(injectedData));

  return base64EncodedData;
}

function encodeUnicodeToBase64(value) {
  return btoa(unescape(encodeURIComponent(value)));
}
