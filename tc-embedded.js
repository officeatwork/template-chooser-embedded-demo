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
    $status.title = status;
    toggleSpinner(false);
  }

  if (event.data && event.data.type === "template-chooser-document-created") {
    const { blob, fileName } = event.data;

    blobDocument = blob;
    $resultFile.innerHTML = fileName;
    toggleSpinner(false);
  }

  if (event.data && event.data.type === "template-chooser-document-uploaded") {
    const { redirectUrl } = event.data;

    let status = "Document was uploaded";
    if (redirectUrl) {
      status += `. Response data: redirectUrl=<a href='${redirectUrl}' target='_blank'>${redirectUrl}</a>`;
    }

    $status.innerHTML = status;
    $status.title = status;
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
  $tcInputUrl.value = defaultTemplateChooserDomain;

  $reload.addEventListener("click", () => {
    reloadIframe();
  });

  $copyUrl.addEventListener("click", () => {
    const embeddedUrl = document.getElementById('template-chooser-embedded-iframe').src;
    navigator.clipboard.writeText(embeddedUrl);
  });

  $resultFile.addEventListener("click", () => {
    saveAs(blobDocument, $resultFile.innerHTML);
  });

  function reloadIframe() {
    $status.innerHTML = "";
    $resultFile.innerHTML = "";
    const $iframeContainer = document.getElementById('iframe-container');

    const iframe = document.createElement('iframe');
    iframe.src = buildEmbeddedUrl();
    iframe.id = 'template-chooser-embedded-iframe';
    iframe.sandbox = 'allow-same-origin allow-popups allow-scripts allow-forms';
    $iframeContainer.innerHTML = '';
    $iframeContainer.appendChild(iframe);
  }

  reloadIframe();
}

function toggleSpinner(show) {
  show ? ($spinner.style.display = "block") : ($spinner.style.display = "none");
}

function buildEmbeddedUrl() {
  const templateUrl = $template.value.trim();
  const base64Template = templateUrl.includes('template=')
    ? new URL(templateUrl).searchParams.get("template")
    : templateUrl || "";
  const uploadUrl = $uploadUrl.value.trim();

  const params = new URLSearchParams(
    `?uploadUrl=${uploadUrl}&template=${base64Template}`
  ).toString();

  const userInputUrl = $tcInputUrl.value || defaultTemplateChooserDomain;
  const [baseUrl, ...routing] = userInputUrl.split("#");

  let url = `${baseUrl}?${params}`;

  if (routing) {
    url += `#${routing.join("#")}`;
  }

  return url;
}
